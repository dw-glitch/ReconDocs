import { canonicalId, norm, text } from "./analysis-engine.js";

export const CROSS_ROLE_LABELS = {
  general: "Base de referência",
  planned: "Base de alocação",
  extra: "Planilha comum",
};

export const CROSS_ROLE_HINTS = {
  general: "A planilha usada como referência da comparação. Opcional.",
  planned: "Estar nesta planilha significa Alocado = Sim; não estar, Alocado = Não. Opcional.",
  extra: "Entra no cruzamento como qualquer outra planilha.",
};

export const MATCH_MODES = {
  smart: "Inteligente (ignora extensão e revisão no fim do código)",
  exact: "Exata (compara o texto normalizado)",
};

/**
 * Chave de cruzamento do documento.
 * `smart` reaproveita a normalização do ReconDocs (remove caminho, extensão e
 * revisão colada ao código); `exact` apenas normaliza acentos, espaços e caixa.
 */
export function crossKey(value, matchMode = "smart") {
  return matchMode === "exact" ? norm(value) : canonicalId(value);
}

function uniqueTexts(values) {
  return [...new Set(values.map((value) => text(value)).filter(Boolean))];
}

function mergeExtras(records) {
  const merged = new Map();
  for (const record of records) {
    for (const [label, value] of Object.entries(record.extras || {})) {
      const content = text(value);
      if (!content) continue;
      if (!merged.has(label)) merged.set(label, new Set());
      merged.get(label).add(content);
    }
  }
  return [...merged.entries()].map(([label, values]) => ({ label, value: [...values].join(" | ") }));
}

function summarizeSource(records) {
  const documents = uniqueTexts(records.map((record) => record.document));
  const statuses = uniqueTexts(records.map((record) => record.status));
  const dates = uniqueTexts(records.map((record) => record.date));
  const extras = mergeExtras(records);
  return {
    present: true,
    count: records.length,
    documents,
    document: documents[0] || "",
    statuses,
    status: statuses.join(" | "),
    dates,
    date: dates[0] || "",
    dateValue: records.map((record) => record.dateValue).find((value) => value instanceof Date) || null,
    extras,
    extrasText: extras.map((extra) => `${extra.label}: ${extra.value}`).join("; "),
    rows: records.map((record) => record.rowNumber).filter((row) => Number.isFinite(row)),
    duplicated: records.length > 1,
    ambiguous: documents.length > 1 || statuses.length > 1,
  };
}

const ABSENT_SOURCE = {
  present: false,
  count: 0,
  documents: [],
  document: "",
  statuses: [],
  status: "",
  dates: [],
  date: "",
  dateValue: null,
  extras: [],
  extrasText: "",
  rows: [],
  duplicated: false,
  ambiguous: false,
};

function presenceState(presentCount, totalSources) {
  if (!presentCount) return { kind: "none", label: "Não localizado" };
  if (totalSources > 1 && presentCount === totalSources) return { kind: "all", label: "Em todas as planilhas" };
  if (presentCount === 1) return { kind: "single", label: "Em apenas uma planilha" };
  return { kind: "partial", label: `Em ${presentCount} de ${totalSources} planilhas` };
}

/**
 * Situação da linha: um rótulo único, do mais grave para o mais brando, para
 * que dê para filtrar por ele tanto na tela quanto no Excel. As observações
 * continuam trazendo o detalhe completo.
 */
function resolveSituation(row, sources, general, planned) {
  if (row.statusDivergence) return { kind: "divergent", label: "Divergência de status" };
  if (general && !row.existsGeneral) return { kind: "missing_reference", label: `Ausente em ${general.label}` };
  if (planned && !row.existsPlanned) return { kind: "not_allocated", label: "Não alocado" };
  if (sources.length > 1 && row.exclusiveIn) return { kind: "exclusive", label: `Só em ${row.exclusiveIn.label}` };
  if (row.missingIn.length === 1) return { kind: "partial", label: `Ausente em ${row.missingIn[0].label}` };
  if (row.missingIn.length) return { kind: "partial", label: `Ausente em ${row.missingIn.length} planilhas` };
  return { kind: "ok", label: "Sem pendências" };
}

function buildObservations(row, sources, general, planned) {
  const notes = [];
  if (general && !row.existsGeneral) notes.push(`Não localizado em ${general.label}.`);
  if (general && row.existsGeneral && !text(row.generalStatus)) notes.push(`Sem status informado em ${general.label}.`);
  if (planned && !row.existsPlanned) notes.push(`Não consta em ${planned.label} (não alocado).`);
  if (planned && row.existsPlanned && general && !row.existsGeneral) {
    notes.push(`Consta em ${planned.label}, mas não em ${general.label}.`);
  }
  if (row.statusDivergence) {
    notes.push(`Divergência de status: ${row.divergentStatuses.map((entry) => `${entry.label} “${entry.status}”`).join(" × ")}.`);
  }
  if (row.missingIn.length && sources.length > 1) {
    notes.push(`Ausente em: ${row.missingIn.map((source) => source.label).join(", ")}.`);
  }
  const duplicated = sources.filter((source) => row.sources[source.id].duplicated);
  if (duplicated.length) {
    notes.push(`Linhas repetidas em: ${duplicated.map((source) => `${source.label} (${row.sources[source.id].count})`).join(", ")}.`);
  }
  if (!notes.length) notes.push("Sem pendências identificadas.");
  return notes.join(" ");
}

/**
 * Cruza N planilhas já mapeadas.
 *
 * @param sources `[{ id, role, label, records }]` — `role` é `general`,
 *   `planned` ou `extra`; `records` é `[{ document, status, date, extras, rowNumber }]`.
 * @param options `{ matchMode: "smart" | "exact" }`.
 */
export function crossReference(sources = [], options = {}) {
  const matchMode = options.matchMode === "exact" ? "exact" : "smart";
  const list = sources.filter((source) => source && source.id);
  const general = list.find((source) => source.role === "general") || null;
  const planned = list.find((source) => source.role === "planned") || null;
  const others = list.filter((source) => source !== general && source !== planned);

  const groups = new Map();
  for (const source of list) {
    for (const record of source.records || []) {
      const key = crossKey(record.document, matchMode);
      if (!key) continue;
      if (!groups.has(key)) groups.set(key, { display: text(record.document), bySource: new Map() });
      const group = groups.get(key);
      if (!group.bySource.has(source.id)) group.bySource.set(source.id, []);
      group.bySource.get(source.id).push(record);
    }
  }

  const rows = [...groups.entries()].map(([key, group]) => {
    const bySource = {};
    for (const source of list) {
      const records = group.bySource.get(source.id) || [];
      bySource[source.id] = records.length ? summarizeSource(records) : { ...ABSENT_SOURCE };
    }

    const presentIn = list.filter((source) => bySource[source.id].present);
    const missingIn = list.filter((source) => !bySource[source.id].present);
    const document = presentIn.map((source) => bySource[source.id].document).find(Boolean) || group.display;

    const statusEntries = presentIn
      .map((source) => ({ id: source.id, label: source.label, role: source.role, status: bySource[source.id].status }))
      .filter((entry) => text(entry.status));
    const distinctStatuses = [...new Set(statusEntries.map((entry) => norm(entry.status)))];
    const statusDivergence = distinctStatuses.length > 1;

    const existsGeneral = Boolean(general && bySource[general.id].present);
    const existsPlanned = Boolean(planned && bySource[planned.id].present);
    const allocation = !planned
      ? { kind: "unknown", label: "—" }
      : existsPlanned
        ? { kind: "allocated", label: "Sim" }
        : { kind: "not_allocated", label: "Não" };

    const row = {
      key,
      document,
      sources: bySource,
      presentIn,
      missingIn,
      presence: presenceState(presentIn.length, list.length),
      existsGeneral,
      existsPlanned,
      generalStatus: general ? bySource[general.id].status : "",
      generalDate: general ? bySource[general.id].date : "",
      plannedStatus: planned ? bySource[planned.id].status : "",
      allocation,
      allocated: allocation.label,
      statusEntries,
      statusDivergence,
      divergentStatuses: statusDivergence ? statusEntries : [],
      otherStatuses: others
        .filter((source) => bySource[source.id].present && text(bySource[source.id].status))
        .map((source) => `${source.label}: ${bySource[source.id].status}`)
        .join("; "),
      complements: list
        .filter((source) => bySource[source.id].extrasText)
        .map((source) => `${source.label} — ${bySource[source.id].extrasText}`)
        .join(" | "),
      onlyGeneral: Boolean(general && existsGeneral && presentIn.length === 1),
      onlyPlanned: Boolean(planned && existsPlanned && presentIn.length === 1),
      exclusiveIn: presentIn.length === 1 ? presentIn[0] : null,
    };

    row.situation = resolveSituation(row, list, general, planned);
    row.observations = buildObservations(row, list, general, planned);
    return row;
  }).sort((left, right) => left.document.localeCompare(right.document, "pt-BR", { numeric: true }));

  return {
    rows,
    sources: list.map((source) => ({
      id: source.id,
      role: source.role,
      label: source.label,
      total: rows.filter((row) => row.sources[source.id].present).length,
    })),
    generalId: general?.id || "",
    plannedId: planned?.id || "",
    generalLabel: general?.label || "",
    plannedLabel: planned?.label || "",
    matchMode,
    metrics: summarizeCross(rows, list),
  };
}

export function summarizeCross(rows, sources = []) {
  const general = sources.find((source) => source.role === "general") || null;
  const planned = sources.find((source) => source.role === "planned") || null;
  const inGeneral = general ? rows.filter((row) => row.existsGeneral) : [];
  const inPlanned = planned ? rows.filter((row) => row.existsPlanned) : [];

  return {
    total: rows.length,
    sheets: sources.length,
    totalGeneral: inGeneral.length,
    totalPlanned: inPlanned.length,
    inCommon: rows.filter((row) => row.existsGeneral && row.existsPlanned).length,
    onlyGeneral: rows.filter((row) => row.onlyGeneral).length,
    onlyPlanned: rows.filter((row) => row.onlyPlanned).length,
    allocated: planned ? inPlanned.length : 0,
    notAllocated: planned ? rows.length - inPlanned.length : 0,
    divergences: rows.filter((row) => row.statusDivergence).length,
    inAllSheets: sources.length > 1 ? rows.filter((row) => row.presence.kind === "all").length : rows.length,
    partial: rows.filter((row) => ["partial", "single"].includes(row.presence.kind)).length,
    exclusive: rows.filter((row) => row.presence.kind === "single").length,
    missingGeneral: general ? rows.length - inGeneral.length : 0,
    bySituation: rows.reduce((counts, row) => {
      const label = row.situation?.label || "Sem pendências";
      counts[label] = (counts[label] || 0) + 1;
      return counts;
    }, {}),
    perSource: Object.fromEntries(sources.map((source) => [
      source.id,
      rows.filter((row) => row.sources[source.id]?.present).length,
    ])),
  };
}

/**
 * Filtros disponíveis para um cruzamento. Os que dependem de uma base de
 * referência ou de alocação só aparecem quando esses papéis foram atribuídos,
 * e usam o nome que o usuário deu à planilha.
 */
export function crossFilters({ generalLabel = "", plannedLabel = "", sheets = 0 } = {}) {
  const filters = [["all", "Todos"]];
  if (sheets > 1) {
    filters.push(["in_all", "Em todas"], ["partial", "Em algumas"], ["exclusive", "Exclusivos"]);
  }
  if (generalLabel) {
    filters.push(["only_general", `Só ${generalLabel}`], ["missing_general", `Ausentes em ${generalLabel}`]);
  }
  if (plannedLabel) {
    filters.push(["only_planned", `Só ${plannedLabel}`], ["not_allocated", "Não alocados"]);
  }
  filters.push(["divergent", "Divergências"]);
  return filters;
}

export function filterCrossRows(rows, filter = "all", query = "") {
  const search = norm(query);
  return rows.filter((row) => {
    const matchesFilter = filter === "all"
      || (filter === "in_all" && row.presence.kind === "all")
      || (filter === "partial" && ["partial", "single"].includes(row.presence.kind))
      || (filter === "exclusive" && row.presence.kind === "single")
      || (filter === "only_general" && row.onlyGeneral)
      || (filter === "only_planned" && row.onlyPlanned)
      || (filter === "not_allocated" && row.allocation.kind === "not_allocated")
      || (filter === "missing_general" && !row.existsGeneral)
      || (filter === "divergent" && row.statusDivergence);
    if (!matchesFilter) return false;
    if (!search) return true;
    return norm([
      row.document,
      row.key,
      row.generalStatus,
      row.otherStatuses,
      row.presence.label,
      row.situation.label,
      row.allocated,
      row.observations,
      row.complements,
    ].join(" ")).includes(search);
  });
}
