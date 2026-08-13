"use client";

import { Fragment, useDeferredValue, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import Link from "next/link";
import * as XLSX from "xlsx";
import {
  AlertTriangle,
  ArrowLeftRight,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleCheck,
  CircleHelp,
  Download,
  FileSpreadsheet,
  Filter,
  Layers,
  LoaderCircle,
  Plus,
  RotateCcw,
  Search,
  Settings2,
  UploadCloud,
  X,
} from "lucide-react";
import { countMappedRecords, labelFromFileName, profileRows, recordFromRow, sampleColumnValues, spreadsheetColumnLetter, uniqueLabelAmong } from "../lib/sheet-profile";
import { readMappingPreset, writeMappingPreset } from "../lib/mapping-presets";
import { CROSS_ROLE_HINTS, CROSS_ROLE_LABELS, crossFilters, crossReference, filterCrossRows, MATCH_MODES, overlapSegments } from "../lib/crosscheck-engine";
import { createBrandLogoDataUrl } from "../lib/report-branding";
import { buildCrossWorkbook } from "../lib/cross-report";

type CrossRole = "general" | "planned" | "extra";
type MatchMode = "smart" | "exact";
type CrossField = "document" | "status" | "date";

type CrossMapping = { document: number; status: number; date: number; extras: number[] };

type CrossSheetProfile = {
  name: string;
  rows: unknown[][];
  startRow: number;
  headerIndex: number;
  dataStart: number;
  headers: string[];
  columnLetters: string[];
  mapping: CrossMapping;
  rowCount: number;
};

type CrossFile = {
  fileName: string;
  fileSize: number;
  sheets: CrossSheetProfile[];
  selectedSheet: number;
};

type CrossSlot = { id: string; role: CrossRole; label: string; fallbackLabel: string; renamed?: boolean; file?: CrossFile };

type CrossRecord = {
  document: string;
  status: string;
  date: string;
  extras: Record<string, string>;
  rowNumber: number;
};

type SourceEntry = {
  present: boolean;
  count: number;
  documents: string[];
  document: string;
  statuses: string[];
  status: string;
  dates: string[];
  date: string;
  extras: { label: string; value: string }[];
  extrasText: string;
  rows: number[];
  duplicated: boolean;
  ambiguous: boolean;
};

type CrossSourceRef = { id: string; role: CrossRole; label: string };

type CrossRow = {
  key: string;
  document: string;
  sources: Record<string, SourceEntry>;
  presentIn: CrossSourceRef[];
  missingIn: CrossSourceRef[];
  presence: { kind: string; label: string };
  existsGeneral: boolean;
  existsPlanned: boolean;
  generalStatus: string;
  generalDate: string;
  plannedStatus: string;
  allocation: { kind: string; label: string };
  allocated: string;
  statusEntries: { id: string; label: string; role: CrossRole; status: string }[];
  statusDivergence: boolean;
  divergentStatuses: { id: string; label: string; status: string }[];
  fieldDivergences: { field: string; entries: { id: string; label: string; value: string }[] }[];
  hasDivergence: boolean;
  otherStatuses: string;
  complements: string;
  onlyGeneral: boolean;
  onlyPlanned: boolean;
  exclusiveIn: CrossSourceRef | null;
  situation: { kind: string; label: string };
  observations: string;
};

type CrossMetrics = {
  total: number;
  sheets: number;
  totalGeneral: number;
  totalPlanned: number;
  inCommon: number;
  onlyGeneral: number;
  onlyPlanned: number;
  allocated: number;
  notAllocated: number;
  divergences: number;
  inAllSheets: number;
  partial: number;
  missingGeneral: number;
  exclusive: number;
  bySituation: Record<string, number>;
  perSource: Record<string, number>;
};

type CrossOutput = {
  rows: CrossRow[];
  sources: (CrossSourceRef & { total: number })[];
  generalId: string;
  plannedId: string;
  generalLabel: string;
  plannedLabel: string;
  matchMode: MatchMode;
  metrics: CrossMetrics;
};

type Progress = { value: number; label: string };

type TableColumn = {
  key: string;
  header: string;
  title?: string;
  width: number;
  cell: (row: CrossRow) => ReactNode;
};

const ROLE_HINTS = CROSS_ROLE_HINTS as Record<CrossRole, string>;
const ROLE_LABELS = CROSS_ROLE_LABELS as Record<CrossRole, string>;

const FIELD_LABELS: Record<CrossField, string> = {
  document: "Coluna do documento",
  status: "Coluna de status",
  date: "Coluna de data",
};

const PAGE_SIZE = 50;

function formatNumber(value: number) {
  return new Intl.NumberFormat("pt-BR").format(value);
}

function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1).replace(".", ",")} MB`;
}

function frame() {
  return new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
}

async function readCrossFile(file: File): Promise<CrossFile> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: false });
  const sheets: CrossSheetProfile[] = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: true,
      blankrows: true,
    }) as unknown[][];
    const profile = profileRows(rows, { columnLabel: (index: number) => `Coluna ${spreadsheetColumnLetter(index, range.s.c)}` });
    const preset = typeof window !== "undefined" ? readMappingPreset(window.localStorage, profile.headers) : null;
    const mapping = preset ? { ...profile.mapping, ...preset } : profile.mapping;
    const columnLetters = profile.headers.map((_, index) => spreadsheetColumnLetter(index, range.s.c));
    return { name, rows, startRow: range.s.r, ...profile, mapping, columnLetters } as CrossSheetProfile;
  });
  const withDocument = sheets.findIndex((sheet) => sheet.mapping.document >= 0 && sheet.rowCount > 0);
  return { fileName: file.name, fileSize: file.size, sheets, selectedSheet: Math.max(0, withDocument) };
}

async function recordsFromSlot(slot: CrossSlot, onProgress?: (done: number, total: number) => void) {
  const sheet = slot.file?.sheets[slot.file.selectedSheet];
  if (!sheet || sheet.mapping.document < 0) return [];
  const records: CrossRecord[] = [];
  for (let index = sheet.dataStart; index < sheet.rows.length; index += 1) {
    const record = recordFromRow(sheet, sheet.rows[index] || [], index) as CrossRecord | null;
    if (record) records.push(record);
    if (index > 0 && index % 5000 === 0) {
      onProgress?.(index, sheet.rows.length);
      await frame();
    }
  }
  return records;
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

async function exportCrossWorkbook(output: CrossOutput) {
  const workbook = buildCrossWorkbook(output, { logoBase64: createBrandLogoDataUrl("CRUZAMENTO DE PLANILHAS") });
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `ReconDocs_Cruzamento_${new Date().toISOString().slice(0, 10)}.xlsx`,
  );
}

function MappingSelect({ sheet, field, onChange }: { sheet: CrossSheetProfile; field: CrossField; onChange: (index: number) => void }) {
  const column = sheet.mapping[field];
  const samples = useMemo(
    () => column >= 0 ? sampleColumnValues(sheet.rows, column, sheet.dataStart) : [],
    [sheet.rows, sheet.dataStart, column],
  );
  return (
    <label className="mapping-field">
      <span>{FIELD_LABELS[field]}{field === "document" ? " *" : ""}</span>
      <select value={column} onChange={(event) => onChange(Number(event.target.value))}>
        <option value={-1}>Não usar</option>
        {sheet.columnLetters.map((letter, index) => (
          <option key={letter} value={index} title={sheet.headers[index]}>Coluna {letter}</option>
        ))}
      </select>
      {column >= 0 && sheet.headers[column] !== `Coluna ${sheet.columnLetters[column]}` && (
        <small className="mapping-sample">Cabeçalho: {sheet.headers[column]}</small>
      )}
      {samples.length > 0 && <small className="mapping-sample">Exemplos: {samples.join(" · ")}</small>}
    </label>
  );
}

function SlotCard({ slot, index, busy, removable, onUpload, onRemoveFile, onRemoveSlot, onSheet, onMapping, onExtras, onLabel, onRole }: {
  slot: CrossSlot;
  index: number;
  busy: boolean;
  removable: boolean;
  onUpload: (file: File) => void;
  onRemoveFile: () => void;
  onRemoveSlot: () => void;
  onSheet: (sheet: number) => void;
  onMapping: (field: CrossField, column: number) => void;
  onExtras: (columns: number[]) => void;
  onLabel: (label: string) => void;
  onRole: (role: CrossRole) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const [panelOpen, setPanelOpen] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const sheet = slot.file?.sheets[slot.file.selectedSheet];
  const documentCount = useMemo(() => sheet ? countMappedRecords(sheet) : 0, [sheet]);
  const detectedState = !sheet || sheet.mapping.document < 0
    ? "unmapped"
    : documentCount === 0
      ? "empty"
      : "ok";
  // A letra identifica a coluna sem ambiguidade; o cabeçalho só aparece
  // junto quando é um cabeçalho de verdade, não o rótulo de reserva "Coluna X".
  const columnLabel = (columnIndex: number) => {
    const letter = sheet!.columnLetters[columnIndex];
    const header = sheet!.headers[columnIndex];
    return header === `Coluna ${letter}` ? `Coluna ${letter}` : `Coluna ${letter} (${header})`;
  };

  return (
    <article className={`upload-card ${slot.file ? "has-file" : ""}`}>
      <div className="upload-card-heading">
        <span className="step-number">{String(index + 1).padStart(2, "0")}</span>
        <h3 className="slot-title" title={slot.label}>{slot.label}</h3>
        {removable && <button type="button" className="icon-button slot-remove" onClick={onRemoveSlot} aria-label={`Remover ${slot.label}`}><X size={16} /></button>}
      </div>
      <p className="slot-hint">{ROLE_HINTS[slot.role]}</p>

      {!slot.file ? (
        <button
          type="button"
          className={`dropzone ${dragging ? "is-dragging" : ""}`}
          onClick={() => input.current?.click()}
          onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={(event) => { event.preventDefault(); setDragging(false); const file = event.dataTransfer.files[0]; if (file) onUpload(file); }}
          disabled={busy}
        >
          {busy ? <LoaderCircle size={23} className="spin" /> : <UploadCloud size={23} strokeWidth={1.8} />}
          <span>Arraste ou <strong>selecione a planilha</strong></span>
        </button>
      ) : (
        <div className="file-loaded">
          <div className="file-summary">
            <span className="file-icon"><FileSpreadsheet size={22} /></span>
            <div>
              <strong title={slot.file.fileName}>{slot.file.fileName}</strong>
              <span>{formatBytes(slot.file.fileSize)} · {formatNumber(sheet?.rowCount || 0)} linhas</span>
            </div>
            <button type="button" className="icon-button" onClick={onRemoveFile} aria-label={`Remover arquivo de ${slot.label}`}><X size={17} /></button>
          </div>
          <div className={`detected-line ${detectedState === "ok" ? "" : "detected-warning"}`}>
            {detectedState === "ok" ? <CircleCheck size={16} /> : <AlertTriangle size={16} />}
            <span>
              {detectedState === "unmapped" && "Selecione a coluna do documento"}
              {detectedState === "empty" && `A ${columnLabel(sheet!.mapping.document)} não trouxe nenhum documento. Verifique o mapeamento.`}
              {detectedState === "ok" && `Documento: ${columnLabel(sheet!.mapping.document)}${sheet!.mapping.status >= 0 ? ` · Status: ${columnLabel(sheet!.mapping.status)}` : ""} · ${formatNumber(documentCount)} encontrado(s)`}
            </span>
          </div>
          <div className="slot-identity">
            <label className="mapping-field">
              <span>Nome no relatório</span>
              <input value={slot.label} onChange={(event) => onLabel(event.target.value)} />
            </label>
            <label className="mapping-field">
              <span>Papel no cruzamento</span>
              <select value={slot.role} onChange={(event) => onRole(event.target.value as CrossRole)}>
                {(["extra", "general", "planned"] as CrossRole[]).map((role) => (
                  <option key={role} value={role}>{ROLE_LABELS[role]}</option>
                ))}
              </select>
            </label>
          </div>
          <button type="button" className="mapping-toggle" onClick={() => setPanelOpen((open) => !open)}>
            <Settings2 size={15} /> Mapear colunas <ChevronDown size={15} className={panelOpen ? "rotate" : ""} />
          </button>
          {panelOpen && sheet && (
            <div className="mapping-panel">
              {slot.file.sheets.length > 1 && (
                <label className="mapping-field">
                  <span>Aba analisada</span>
                  <select value={slot.file.selectedSheet} onChange={(event) => onSheet(Number(event.target.value))}>
                    {slot.file.sheets.map((item, sheetIndex) => (
                      <option key={`${item.name}-${sheetIndex}`} value={sheetIndex}>{item.name} · {formatNumber(item.rowCount)} linhas</option>
                    ))}
                  </select>
                </label>
              )}
              {(["document", "status", "date"] as CrossField[]).map((field) => (
                <MappingSelect key={field} sheet={sheet} field={field} onChange={(column) => onMapping(field, column)} />
              ))}
              <div className="mapping-field">
                <span>Outras colunas relevantes</span>
                <div className="extras-list">
                  {sheet.headers.map((header, column) => (
                    <label key={`${header}-${column}`} className="extra-option">
                      <input
                        type="checkbox"
                        checked={sheet.mapping.extras.includes(column)}
                        onChange={(event) => onExtras(event.target.checked
                          ? [...sheet.mapping.extras, column].sort((left, right) => left - right)
                          : sheet.mapping.extras.filter((item) => item !== column))}
                      />
                      <span>{header}</span>
                    </label>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>
      )}
      <input
        ref={input}
        type="file"
        accept=".xlsx,.xls,.xlsm,.csv,.tsv"
        hidden
        onChange={(event) => { const file = event.target.files?.[0]; if (file) onUpload(file); event.target.value = ""; }}
      />
    </article>
  );
}

type OverlapSegment = { key: string; label: string; count: number; percent: number };

/**
 * Resumo visual de sobreposição: em quantas planilhas cada documento
 * aparece, numa barra empilhada que responde em um olhar o que hoje exigiria
 * ler vários números separados.
 */
function OverlapBar({ segments }: { segments: OverlapSegment[] }) {
  const visible = segments.filter((segment) => segment.count > 0);
  if (!visible.length) return null;
  return (
    <div className="overlap-bar-wrap">
      <div className="overlap-bar" role="img" aria-label={visible.map((segment) => `${segment.label}: ${formatNumber(segment.count)}`).join(", ")}>
        {visible.map((segment) => (
          <span key={segment.key} className={`overlap-segment overlap-${segment.key}`} style={{ width: `${Math.max(segment.percent, visible.length > 1 ? 2 : 100)}%` }} />
        ))}
      </div>
      <div className="overlap-legend">
        {segments.map((segment) => (
          <div key={segment.key} className="overlap-legend-item">
            <span className={`overlap-dot overlap-${segment.key}`} />
            <b>{formatNumber(segment.count)}</b>
            <span>{segment.label}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function DetailPanel({ row, sources }: { row: CrossRow; sources: (CrossSourceRef & { total: number })[] }) {
  return (
    <div className="detail-panel">
      {sources.map((source) => {
        const entry = row.sources[source.id];
        return (
          <div key={source.id} className="detail-block">
            <span className="detail-label">{source.label}</span>
            <strong>{entry?.present ? "Encontrado" : "Não encontrado"}</strong>
            {entry?.present && (
              <>
                <p><b>Documento:</b> {entry.documents.join(" | ")}</p>
                <p><b>Status:</b> {entry.status || "vazio"}</p>
                {entry.date && <p><b>Data:</b> {entry.date}</p>}
                {entry.extrasText && <p><b>Complementos:</b> {entry.extrasText}</p>}
                <small>Linha(s) de origem: {entry.rows.join(", ") || "—"}</small>
              </>
            )}
          </div>
        );
      })}
      <div className="detail-block detail-wide">
        <span className="detail-label">Situação</span>
        <strong className={`situation-text situation-${row.situation.kind}`}>{row.situation.label}</strong>
        <p>{row.observations}</p>
      </div>
    </div>
  );
}

export default function CrossCheckPage() {
  const [slots, setSlots] = useState<CrossSlot[]>([
    { id: "sheet-1", role: "extra", label: "Planilha 1", fallbackLabel: "Planilha 1" },
    { id: "sheet-2", role: "extra", label: "Planilha 2", fallbackLabel: "Planilha 2" },
  ]);
  const [busySlot, setBusySlot] = useState<string | null>(null);
  const [matchMode, setMatchMode] = useState<MatchMode>("smart");
  const [error, setError] = useState("");
  const [progress, setProgress] = useState<Progress | null>(null);
  const [output, setOutput] = useState<CrossOutput | null>(null);
  const [activeFilter, setActiveFilter] = useState("all");
  const [query, setQuery] = useState("");
  const deferredQuery = useDeferredValue(query);
  const [page, setPage] = useState(1);
  const [expanded, setExpanded] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);

  const loadedSlots = slots.filter((slot) => slot.file && (slot.file.sheets[slot.file.selectedSheet]?.mapping.document ?? -1) >= 0);
  const ready = loadedSlots.length >= 2;

  const filtered = useMemo<CrossRow[]>(
    () => output ? filterCrossRows(output.rows, activeFilter, deferredQuery) as CrossRow[] : [],
    [output, activeFilter, deferredQuery],
  );
  const pageCount = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const visibleRows = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const invalidate = () => { setOutput(null); setExpanded(null); };

  const updateSlot = (id: string, transform: (slot: CrossSlot) => CrossSlot) => {
    setSlots((current) => current.map((slot) => slot.id === id ? transform(slot) : slot));
    invalidate();
  };

  const updateSheet = (id: string, transform: (sheet: CrossSheetProfile) => CrossSheetProfile) => {
    updateSlot(id, (slot) => slot.file
      ? {
        ...slot,
        file: {
          ...slot.file,
          sheets: slot.file.sheets.map((sheet, index) => index === slot.file!.selectedSheet ? transform(sheet) : sheet),
        },
      }
      : slot);
  };

  const updateSheetMapping = (id: string, transform: (sheet: CrossSheetProfile) => CrossSheetProfile) => {
    updateSheet(id, (sheet) => {
      const updated = transform(sheet);
      if (typeof window !== "undefined") writeMappingPreset(window.localStorage, updated.headers, updated.mapping);
      return updated;
    });
  };

  const handleFile = async (id: string, file: File) => {
    setError("");
    if (!/\.(xlsx|xls|xlsm|csv|tsv)$/i.test(file.name)) {
      setError("Use uma planilha .xlsx, .xls, .xlsm, .csv ou .tsv.");
      return;
    }
    setBusySlot(id);
    try {
      await frame();
      const parsed = await readCrossFile(file);
      if (!parsed.sheets.length) throw new Error("A planilha não possui abas legíveis.");
      setSlots((current) => current.map((slot) => {
        if (slot.id !== id) return slot;
        const fromFile = labelFromFileName(file.name);
        const taken = current.filter((item) => item.id !== id).map((item) => item.label);
        const label = slot.renamed || !fromFile ? slot.label : uniqueLabelAmong(fromFile, taken);
        return { ...slot, file: parsed, label };
      }));
      invalidate();
    } catch (cause) {
      setError(cause instanceof Error ? `Não foi possível ler ${file.name}: ${cause.message}` : `Não foi possível ler ${file.name}.`);
    } finally {
      setBusySlot(null);
    }
  };

  const addSlot = () => {
    setSlots((current) => {
      const nextIndex = current.length + 1;
      const label = `Planilha ${nextIndex}`;
      return [...current, { id: `sheet-${Date.now()}`, role: "extra", label, fallbackLabel: label }];
    });
    invalidate();
  };

  const runCross = async () => {
    if (!ready) return;
    setError("");
    setOutput(null);
    setProgress({ value: 5, label: "Preparando as planilhas" });
    try {
      const sources = [];
      for (let index = 0; index < loadedSlots.length; index += 1) {
        const slot = loadedSlots[index];
        const base = 8 + Math.round((index / loadedSlots.length) * 78);
        setProgress({ value: base, label: `Lendo ${slot.label}` });
        const records = await recordsFromSlot(slot, (done, total) => {
          setProgress({
            value: base + Math.round((done / Math.max(1, total)) * (70 / loadedSlots.length)),
            label: `Lendo ${slot.label} · ${formatNumber(done)} linhas`,
          });
        });
        sources.push({ id: slot.id, role: slot.role, label: slot.label.trim() || `Planilha ${index + 1}`, records });
      }
      setProgress({ value: 90, label: "Cruzando os documentos" });
      await frame();
      const result = crossReference(sources, { matchMode }) as unknown as CrossOutput;
      setProgress({ value: 100, label: "Cruzamento concluído" });
      setOutput(result);
      setActiveFilter("all");
      setQuery("");
      setPage(1);
      setExpanded(null);
      setTimeout(() => setProgress(null), 450);
      requestAnimationFrame(() => document.getElementById("cross-results")?.scrollIntoView({ behavior: "smooth", block: "start" }));
    } catch (cause) {
      setProgress(null);
      setError(cause instanceof Error ? cause.message : "O cruzamento não pôde ser concluído.");
    }
  };

  const exportReport = async () => {
    if (!output) return;
    setExporting(true);
    setError("");
    try {
      await exportCrossWorkbook(output);
    } catch (cause) {
      setError(cause instanceof Error ? `Não foi possível exportar: ${cause.message}` : "Não foi possível exportar o relatório.");
    } finally {
      setExporting(false);
    }
  };

  const reset = () => {
    setSlots([
      { id: "sheet-1", role: "extra", label: "Planilha 1", fallbackLabel: "Planilha 1" },
      { id: "sheet-2", role: "extra", label: "Planilha 2", fallbackLabel: "Planilha 2" },
    ]);
    setOutput(null);
    setError("");
    setProgress(null);
    setQuery("");
    setPage(1);
    setExpanded(null);
  };

  const metrics = output?.metrics;
  const generalLabel = output?.generalLabel || "";
  const plannedLabel = output?.plannedLabel || "";
  const filters = useMemo(
    () => crossFilters({ generalLabel, plannedLabel, sheets: output?.sources.length || 0 }) as [string, string][],
    [generalLabel, plannedLabel, output],
  );
  // Uma única passada pelas linhas por filtro, memoizada — em vez de recontar
  // a cada tecla digitada na busca, já que a contagem dos rótulos não depende
  // do texto buscado.
  const filterCounts = useMemo(() => {
    if (!output) return {} as Record<string, number>;
    return Object.fromEntries(filters.map(([key]) => [key, filterCrossRows(output.rows, key, "").length]));
  }, [output, filters]);
  const filterCount = (filter: string) => filterCounts[filter] ?? 0;

  // As colunas da tabela seguem as planilhas carregadas: as de referência e de
  // alocação só existem quando esses papéis foram atribuídos.
  const tableColumns = useMemo<TableColumn[]>(() => {
    const columns: TableColumn[] = [
      {
        key: "document",
        header: "Documento",
        width: generalLabel || plannedLabel ? 30 : 36,
        cell: (row) => (
          <>
            <strong className="document-code">{row.document}</strong>
            <small>{row.presence.label} · {row.presentIn.map((source) => source.label).join(" · ") || "Sem fonte"}</small>
          </>
        ),
      },
      {
        key: "situation",
        header: "Situação",
        width: 20,
        cell: (row) => <span className={`status-badge situation-${row.situation.kind}`}>{row.situation.label}</span>,
      },
    ];

    if (generalLabel) {
      columns.push({
        key: "general",
        header: generalLabel,
        title: `Existe em ${generalLabel}`,
        width: 15,
        cell: (row) => (
          <>
            <span className={`status-badge status-${row.existsGeneral ? "allocated" : "not_allocated"}`}>{row.existsGeneral ? "Sim" : "Não"}</span>
            {row.generalStatus && <small className="cell-note">{row.generalStatus}</small>}
          </>
        ),
      });
    }

    if (plannedLabel) {
      columns.push({
        key: "allocated",
        header: "Alocado",
        title: `Presente em ${plannedLabel}`,
        width: 12,
        cell: (row) => <span className={`status-badge status-${row.allocation.kind}`}>{row.allocated}</span>,
      });
    }

    columns.push({
      key: "status",
      header: "Status por planilha",
      width: generalLabel || plannedLabel ? 22 : 43,
      cell: (row) => {
        if (row.statusDivergence) {
          return <span className="difference-text status-divergence">{row.divergentStatuses.map((entry) => `${entry.label}: ${entry.status}`).join(" × ")}</span>;
        }
        if (row.fieldDivergences.length) {
          const first = row.fieldDivergences[0];
          return (
            <span className="difference-text status-divergence">
              {first.field}: {first.entries.map((entry) => `${entry.label}: ${entry.value}`).join(" × ")}
              {row.fieldDivergences.length > 1 && ` (+${row.fieldDivergences.length - 1})`}
            </span>
          );
        }
        return <span className="nt-result">{row.statusEntries.map((entry) => `${entry.label}: ${entry.status}`)[0] || "Sem status informado"}</span>;
      },
    });

    columns.push({
      key: "chevron",
      header: "",
      width: 5,
      cell: (row) => <ChevronDown size={17} className={expanded === row.key ? "rotate" : ""} />,
    });

    return columns;
  }, [generalLabel, plannedLabel, expanded]);

  return (
    <main>
      <header className="topbar">
        <Link className="brand" href="/" aria-label="ReconDocs — início">
          <span className="brand-mark">
            <ArrowLeftRight size={21} />
            <span className="brand-check"><Check size={10} strokeWidth={3.2} /></span>
          </span>
          <span className="brand-copy"><b><span>RECON</span><em>DOCS</em></b><small>CRUZAMENTO DE PLANILHAS</small></span>
        </Link>
        <nav className="module-nav">
          <Link href="/">Conferência SGP × SIGEM</Link>
          <Link href="/cruzamento" className="active" aria-current="page">Cruzamento de planilhas</Link>
          {slots.some((slot) => slot.file) && <button type="button" className="quiet-button" onClick={reset}><RotateCcw size={15} /> Limpar</button>}
        </nav>
      </header>

      <section className="workspace-section" id="top" aria-labelledby="cross-title">
        <div className="section-heading">
          <div>
            <h1 id="cross-title">Cruzamento de planilhas</h1>
            <p className="section-subtitle">
              Carregue quantas planilhas quiser, de qualquer origem e layout, e cruze todas de uma vez. O módulo detecta as colunas
              sozinho e você ajusta o que quiser. Opcionalmente, marque uma planilha como base de referência e outra como base de
              alocação para ativar as colunas de <b>Existe</b> e <b>Alocado</b>.
            </p>
          </div>
          <span className="upload-count">{loadedSlots.length} de {slots.length} prontas</span>
        </div>

        <div className="upload-grid cross-slots">
          {slots.map((slot, index) => (
            <SlotCard
              key={slot.id}
              slot={slot}
              index={index}
              busy={busySlot === slot.id}
              removable={slots.length > 1}
              onUpload={(file) => handleFile(slot.id, file)}
              onRemoveFile={() => updateSlot(slot.id, (item) => ({ ...item, file: undefined, label: item.renamed ? item.label : item.fallbackLabel }))}
              onRemoveSlot={() => { setSlots((current) => current.filter((item) => item.id !== slot.id)); invalidate(); }}
              onSheet={(selectedSheet) => updateSlot(slot.id, (item) => item.file ? { ...item, file: { ...item.file, selectedSheet } } : item)}
              onMapping={(field, column) => updateSheetMapping(slot.id, (sheet) => ({ ...sheet, mapping: { ...sheet.mapping, [field]: column } }))}
              onExtras={(extras) => updateSheetMapping(slot.id, (sheet) => ({ ...sheet, mapping: { ...sheet.mapping, extras } }))}
              onLabel={(label) => updateSlot(slot.id, (item) => ({ ...item, label, renamed: true }))}
              onRole={(role) => updateSlot(slot.id, (item) => ({ ...item, role }))}
            />
          ))}
          <button type="button" className="add-slot" onClick={addSlot}>
            <Plus size={20} />
            <b>Adicionar planilha</b>
            <span>Sem limite de quantidade</span>
          </button>
        </div>

        <div className="control-panel">
          <div className="scope-group">
            <b className="control-label">Comparação dos documentos</b>
            <div className="scope-options">
              {(Object.keys(MATCH_MODES) as MatchMode[]).map((mode) => (
                <button
                  key={mode}
                  type="button"
                  className={matchMode === mode ? "active" : ""}
                  aria-pressed={matchMode === mode}
                  onClick={() => { setMatchMode(mode); invalidate(); }}
                >
                  <b>{mode === "smart" ? "Inteligente" : "Exata"}</b>
                </button>
              ))}
            </div>
            <p className="control-hint">{MATCH_MODES[matchMode]}</p>
          </div>
          <button type="button" className="primary-button" onClick={runCross} disabled={!ready || Boolean(progress) || Boolean(busySlot)}>
            {progress ? <LoaderCircle size={18} className="spin" /> : <Layers size={18} />}
            {progress ? progress.label : "Cruzar planilhas"}
          </button>
        </div>

        {!ready && <p className="control-hint control-hint-block">Carregue pelo menos duas planilhas com a coluna do documento mapeada para iniciar o cruzamento. Não há limite de planilhas.</p>}
        {error && <div className="error-banner" role="alert"><AlertTriangle size={18} /><span>{error}</span></div>}
        {progress && <div className="progress-track" aria-label={progress.label}><span style={{ width: `${progress.value}%` }} /></div>}
      </section>

      {output && metrics && (
        <section className="results-section" id="cross-results">
          <div className="results-heading">
            <div>
              <h2>Resultado do cruzamento</h2>
              <p>{formatNumber(metrics.total)} documentos distintos · {output.sources.length} planilhas cruzadas</p>
            </div>
            <button type="button" className="export-button" onClick={exportReport} disabled={exporting}>
              {exporting ? <LoaderCircle size={17} className="spin" /> : <Download size={17} />}
              {exporting ? "Gerando relatório" : "Exportar relatório Excel"}
            </button>
          </div>

          <div className="metrics-grid cross-metrics">
            <article className="metric-card"><span>DOCUMENTOS ANALISADOS</span><strong>{formatNumber(metrics.total)}</strong></article>
            <article className="metric-card metric-good"><span>EM TODAS AS PLANILHAS</span><strong>{formatNumber(metrics.inAllSheets)}</strong></article>
            <article className="metric-card metric-warn"><span>EM APENAS ALGUMAS</span><strong>{formatNumber(metrics.partial)}</strong></article>
            <article className="metric-card metric-warn"><span>EXCLUSIVOS DE UMA</span><strong>{formatNumber(metrics.exclusive)}</strong></article>
            {output.sources.map((source) => (
              <article key={source.id} className="metric-card">
                <span title={source.label}>{source.label.toUpperCase()}</span>
                <strong>{formatNumber(source.total)}</strong>
              </article>
            ))}
            {generalLabel && (
              <article className="metric-card metric-bad">
                <span title={`Ausentes em ${generalLabel}`}>AUSENTES EM {generalLabel.toUpperCase()}</span>
                <strong>{formatNumber(metrics.missingGeneral)}</strong>
              </article>
            )}
            {plannedLabel && <article className="metric-card metric-good"><span>ALOCADOS</span><strong>{formatNumber(metrics.allocated)}</strong></article>}
            {plannedLabel && <article className="metric-card metric-bad"><span>NÃO ALOCADOS</span><strong>{formatNumber(metrics.notAllocated)}</strong></article>}
            {generalLabel && plannedLabel && <article className="metric-card metric-good"><span>EM COMUM</span><strong>{formatNumber(metrics.inCommon)}</strong></article>}
          </div>

          <OverlapBar segments={overlapSegments(metrics)} />

          <div className="insight-strip">
            <div><b>{formatNumber(metrics.divergences)}</b><span>divergências</span></div>
            <div><b>{formatNumber(metrics.sheets)}</b><span>planilhas cruzadas</span></div>
            {generalLabel
              ? <div><b>{formatNumber(metrics.onlyGeneral)}</b><span>só em {generalLabel}</span></div>
              : <div><b>{formatNumber(metrics.exclusive)}</b><span>exclusivos de uma planilha</span></div>}
            {plannedLabel
              ? <div><b>{formatNumber(metrics.onlyPlanned)}</b><span>só em {plannedLabel}</span></div>
              : <div><b>{formatNumber(metrics.partial)}</b><span>ausentes em alguma planilha</span></div>}
          </div>

          <div className="results-card">
            <div className="table-toolbar">
              <div className="filter-tabs" role="tablist" aria-label="Filtros do cruzamento">
                {filters.map(([key, label]: string[]) => (
                  <button key={key} type="button" role="tab" aria-selected={activeFilter === key} className={activeFilter === key ? "active" : ""} onClick={() => { setActiveFilter(key); setPage(1); setExpanded(null); }}>
                    {label}<span>{formatNumber(filterCount(key))}</span>
                  </button>
                ))}
              </div>
              <label className="search-field">
                <Search size={17} />
                <input value={query} onChange={(event) => { setQuery(event.target.value); setPage(1); }} placeholder="Buscar documento" />
                <span className="sr-only">Buscar documento</span>
              </label>
            </div>

            <div className="table-meta">
              <span><Filter size={14} /> Exibindo {formatNumber(filtered.length)} resultados</span>
              <small>Clique em uma linha para ver cada planilha</small>
            </div>

            <div className="table-scroll">
              <table className="cross-table">
                <colgroup>
                  {tableColumns.map((column) => <col key={column.key} style={{ width: `${column.width}%` }} />)}
                </colgroup>
                <thead>
                  <tr>
                    {tableColumns.map((column) => (
                      <th key={column.key} title={column.title || undefined}>{column.header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((row) => (
                    <Fragment key={row.key}>
                      <tr
                        className={expanded === row.key ? "expanded-row" : ""}
                        onClick={() => setExpanded(expanded === row.key ? null : row.key)}
                        tabIndex={0}
                        onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") setExpanded(expanded === row.key ? null : row.key); }}
                      >
                        {tableColumns.map((column) => <td key={column.key}>{column.cell(row)}</td>)}
                      </tr>
                      {expanded === row.key && (
                        <tr className="details-row">
                          <td colSpan={tableColumns.length}><DetailPanel row={row} sources={output.sources} /></td>
                        </tr>
                      )}
                    </Fragment>
                  ))}
                  {!visibleRows.length && (
                    <tr>
                      <td colSpan={tableColumns.length}>
                        <div className="empty-state">
                          <CircleHelp size={28} />
                          <b>Nenhum documento neste filtro</b>
                          <span>Altere o filtro ou a busca para visualizar outros resultados.</span>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="pagination">
              <span>Página {page} de {pageCount}</span>
              <div>
                <button type="button" onClick={() => { setPage((current) => Math.max(1, current - 1)); setExpanded(null); }} disabled={page === 1} aria-label="Página anterior"><ChevronLeft size={17} /></button>
                <button type="button" onClick={() => { setPage((current) => Math.min(pageCount, current + 1)); setExpanded(null); }} disabled={page === pageCount} aria-label="Próxima página"><ChevronRight size={17} /></button>
              </div>
            </div>
          </div>
        </section>
      )}
    </main>
  );
}
