import {
  EAP_CURRENT_TO_LEGACY,
  EAP_LEGACY_TO_CURRENT,
  ET_REPORT_CODES,
  ET_REPORT_SUCCESSORS,
  N1710_CATEGORIES,
  N2064_REVISION_ALPHABET,
} from "./normative-catalog.js";

const CACHE_LIMIT = 250000;
const parseCache = new Map();
const identityCache = new Map();

const KNOWN_FILE_EXTENSION = /\.(?:PDF|DWG|DGN|DOCX?|XLSX?|XLSM|CSV|TSV|ZIP)$/i;
const REVISION_SUFFIX = /(?:\s*[-–—_]?\s*)(?:REV(?:IS(?:ÃO|AO))?\.?|REVISÃO)\s*[:.\-]?\s*([A-Z0-9]{1,5})(?:\s*\(\d+\))?\s*$/i;
const STANDARD_CODE_WITH_REVISION = /^((?:[A-Z]{1,4}-\d{4}\.\d{2}-\d{5}-\d{3}-[A-Z0-9]{3}-\d{3}|\d{4}\.\d{7}\.\d{2}\.\d{1,2}-[A-Z0-9]{3}-CV-[A-Z0-9]{3}-\d{3,4}))(?:\s*[-_–—]\s*([A-Z]{1,3}|\d{1,3})(?:\s*\(\d+\))?)?$/i;
const ET_PATTERN = /^(?<contract>[^_]+)_RNEST_(?<unit>[^_]+)_(?<eap>\d+(?:\.\d+){3})_(?<discipline>[^_]+)_(?<report>[^_]+)_(?<tag>.+)$/i;
const CV_PATTERN = /^(?<installation>\d{4}\.\d{7}\.\d{2}\.\d{1,2})-(?<origin>[A-Z0-9]{2,5})-CV-(?<class>[A-Z0-9]{2,5})-(?<sequence>\d{3,4})$/i;
const N1710_PATTERN = /(?<category>[A-Z]{2})-(?<installation>\d{4}\.\d{2})-(?<area>\d{5})-(?<class>\d{3})-(?<origin>[A-Z0-9]{3})-(?<sequence>\d{3})(?:-(?<language>[A-Z]{2}))?$/i;

function remember(cache, key, value) {
  if (cache.size >= CACHE_LIMIT) cache.clear();
  cache.set(key, value);
  return value;
}

export function text(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return new Intl.DateTimeFormat("pt-BR").format(value);
  return String(value).trim();
}

export function norm(value) {
  return text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[–—]/g, "-")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

export function revisionId(value) {
  return norm(value)
    .replace(/^(?:REV(?:ISAO)?\.?|REVISAO)\s*[:.\-]?\s*/, "")
    .replace(/\s*\(\d+\)\s*$/, "")
    .replace(/^[\s._-]+|[\s._-]+$/g, "")
    .trim();
}

export function parseDocumentReference(value, explicitRevision = "") {
  const original = text(value);
  const fileName = original.split(/[\\/]/).pop() || original;
  const withoutExtension = fileName.replace(KNOWN_FILE_EXTENSION, "").trim();
  const revisionMatch = withoutExtension.match(REVISION_SUFFIX);
  const standardMatch = revisionMatch ? null : withoutExtension.match(STANDARD_CODE_WITH_REVISION);
  const embeddedRevision = revisionMatch
    ? revisionId(revisionMatch[1])
    : standardMatch?.[2]
      ? revisionId(standardMatch[2])
      : "";
  const code = (revisionMatch
    ? withoutExtension.slice(0, revisionMatch.index)
    : standardMatch?.[2]
      ? standardMatch[1]
      : withoutExtension)
    .replace(/[\s\-–—_]+$/g, "")
    .trim();
  const mappedRevision = revisionId(explicitRevision);
  return {
    code,
    revision: mappedRevision || embeddedRevision,
    mappedRevision,
    embeddedRevision,
    revisionConflict: Boolean(mappedRevision && embeddedRevision && mappedRevision !== embeddedRevision),
    original,
  };
}

export function canonicalId(value) {
  return norm(parseDocumentReference(value).code)
    .replace(/\s*([_.-])\s*/g, "$1")
    .replace(/\s+/g, " ");
}

export function displayDocumentCode(value) {
  return text(value).replace(/NT-/gi, "nt-");
}

export function parseEt(value) {
  const canonical = canonicalId(value);
  const match = canonical.match(ET_PATTERN);
  if (!match?.groups) return null;
  const tag = match.groups.tag.replace(/^NT-/i, "");
  const embeddedN1710 = tag.match(N1710_PATTERN)?.[0] || "";
  const report = norm(match.groups.report);
  const eap = match.groups.eap;
  return {
    kind: "et",
    canonical,
    contract: match.groups.contract,
    unit: match.groups.unit,
    eap,
    discipline: match.groups.discipline,
    report,
    tag,
    nt: /^NT-/i.test(match.groups.tag),
    embeddedN1710,
    reportKnown: ET_REPORT_CODES.has(report),
    eapCurrent: EAP_LEGACY_TO_CURRENT.get(eap) || eap,
    eapLegacy: EAP_CURRENT_TO_LEGACY.get(eap) || "",
  };
}

export function parseCv(value) {
  const canonical = canonicalId(value);
  const match = canonical.match(CV_PATTERN);
  return match?.groups ? { kind: "cv", canonical, ...match.groups } : null;
}

export function parseN1710(value) {
  const canonical = canonicalId(value);
  const match = canonical.match(N1710_PATTERN);
  if (!match?.groups) return null;
  return {
    kind: "n1710",
    canonical: match[0],
    ...match.groups,
    categoryKnown: N1710_CATEGORIES.has(match.groups.category),
    wrapped: match[0] !== canonical,
  };
}

export function parseNormativeDocument(value) {
  const key = canonicalId(value);
  if (parseCache.has(key)) return parseCache.get(key);
  const et = parseEt(key);
  const n1710 = parseN1710(key);
  const cv = parseCv(key);
  const parsed = et || cv || n1710 || { kind: "unknown", canonical: key };
  const issues = [];
  if (et && !et.reportKnown) issues.push(`Tipo ET não catalogado: ${et.report}`);
  if (n1710 && !n1710.categoryKnown) issues.push(`Categoria N-1710 não catalogada: ${n1710.category}`);
  parsed.issues = issues;
  return remember(parseCache, key, parsed);
}

export function documentTag(value) {
  const parsed = parseNormativeDocument(value);
  if (parsed.kind === "et") return parsed.embeddedN1710 || parsed.tag;
  return parsed.canonical;
}

export function identityKey(value) {
  const canonical = canonicalId(value);
  if (!canonical) return "";
  if (identityCache.has(canonical)) return identityCache.get(canonical);
  const parsed = parseNormativeDocument(canonical);
  let key = `RAW::${canonical}`;
  if (parsed.kind === "et") key = parsed.embeddedN1710 ? `N1710::${parsed.embeddedN1710}` : `ET::${parsed.tag}`;
  if (parsed.kind === "cv") key = `CV::${parsed.canonical}`;
  if (parsed.kind === "n1710") key = `N1710::${parsed.canonical}`;
  return remember(identityCache, canonical, key);
}

export function documentKind(value) {
  return parseNormativeDocument(value).kind;
}

export function isEtDocument(value) {
  return parseNormativeDocument(value).kind === "et";
}

export function ntNeutralKey(value) {
  const parsed = parseEt(value);
  return parsed ? `${parsed.canonical.slice(0, parsed.canonical.length - parsed.tag.length - (parsed.nt ? 3 : 0))}${parsed.tag}` : canonicalId(value);
}

export function ntPrefixVariant(value) {
  const parsed = parseEt(value);
  if (!parsed) return "";
  const prefix = parsed.canonical.slice(0, parsed.canonical.length - parsed.tag.length - (parsed.nt ? 3 : 0));
  return displayDocumentCode(parsed.nt ? `${prefix}${parsed.tag}` : `${prefix}NT-${parsed.tag}`);
}

export function searchKeys(value) {
  const key = displayDocumentCode(canonicalId(value));
  return key && isEtDocument(key) ? [...new Set([key, ntPrefixVariant(key)].filter(Boolean))] : key ? [key] : [];
}

export function etDocumentSubtype(value, title = "") {
  const parsed = parseEt(value);
  if (!parsed) return "not_et";
  const searchable = norm(`${parsed.report} ${title}`);
  if (/^RIR(?:-|$)/.test(parsed.report) || /\bRIR\b|INSPECAO DE RECEBIMENTO/.test(searchable)) return "rir";
  if (/^(?:C&M|CM|CCM)(?:-|$)/.test(parsed.report) || /\bC\s*&\s*M\b|\bC E M\b|COMPLETACAO MECANICA|COMISSIONAMENTO/.test(searchable)) return "cm";
  return "other";
}

function sameSuccessor(left, right) {
  return ET_REPORT_SUCCESSORS.get(left) === right || ET_REPORT_SUCCESSORS.get(right) === left;
}

export function compareNormativeCodes(leftValue, rightValue) {
  const leftCode = canonicalId(leftValue);
  const rightCode = canonicalId(rightValue);
  const left = parseNormativeDocument(leftCode);
  const right = parseNormativeDocument(rightCode);
  if (!leftCode || !rightCode) return { kind: "missing", label: "Código ausente", detail: "Uma das fontes não possui código." };
  if (leftCode === rightCode) return { kind: "exact", label: "Mesmo código", detail: "Código idêntico nas duas fontes." };
  if (left.kind === "et" && right.kind === "et" && left.tag === right.tag) {
    const differences = [];
    if (left.nt !== right.nt) differences.push("forma com/sem nt-");
    if (left.eap !== right.eap) differences.push(EAP_LEGACY_TO_CURRENT.get(left.eap) === right.eap || EAP_LEGACY_TO_CURRENT.get(right.eap) === left.eap ? "transição de EAP" : `EAP ${left.eap} × ${right.eap}`);
    if (left.report !== right.report) differences.push(sameSuccessor(left.report, right.report) ? `tipo sucessor ${left.report} → ${right.report}` : `tipo ${left.report} × ${right.report}`);
    if (left.discipline !== right.discipline) differences.push(`disciplina ${left.discipline} × ${right.discipline}`);
    const kind = left.nt !== right.nt && differences.length === 1
      ? "nt_variant"
      : differences.some((item) => item === "transição de EAP")
        ? "eap_transition"
        : differences.some((item) => item.startsWith("tipo sucessor"))
          ? "normative_successor"
          : "same_tag";
    return { kind, label: "Mesmo TAG; código alterado", detail: differences.join("; ") || "Mesmo TAG com estrutura de código diferente." };
  }
  if (identityKey(leftCode) === identityKey(rightCode)) {
    return { kind: "same_tag", label: "Mesmo documento; código alterado", detail: "O identificador técnico é o mesmo nas duas fontes." };
  }
  return { kind: "different", label: "Documentos diferentes", detail: `SGP “${displayDocumentCode(leftCode)}” × SIGEM “${displayDocumentCode(rightCode)}”` };
}

function revisionOrdinal(value) {
  const id = revisionId(value);
  if (!id) return null;
  if (/^\d+$/.test(id)) return Number(id);
  if (!/^[A-Z]+$/.test(id) || [...id].some((letter) => !N2064_REVISION_ALPHABET.includes(letter))) return null;
  let ordinal = 0;
  for (const letter of id) ordinal = ordinal * N2064_REVISION_ALPHABET.length + N2064_REVISION_ALPHABET.indexOf(letter) + 1;
  return ordinal;
}

export function revisionInfo(sgpValue, sigemValue) {
  const sgp = revisionId(sgpValue);
  const sigem = revisionId(sigemValue);
  if (!sgp && !sigem) return { kind: "empty", label: "Sem revisão", detail: "Revisão não informada nas duas fontes.", sgp, sigem, n2064: true };
  if (!sgp || !sigem) return { kind: "missing", label: "Revisão ausente", detail: `SGP “${sgp || "vazio"}” × SIGEM “${sigem || "vazio"}”`, sgp, sigem, n2064: false };
  if (sgp === sigem) return { kind: "equal", label: `Revisão ${sgp}`, detail: "Mesma revisão nas duas fontes.", sgp, sigem, n2064: revisionOrdinal(sgp) !== null };
  const left = revisionOrdinal(sgp);
  const right = revisionOrdinal(sigem);
  const direction = left !== null && right !== null ? (right > left ? "SIGEM posterior ao SGP" : "SGP posterior ao SIGEM") : "ordem requer validação";
  return { kind: "different", label: `Revisão ${sgp} × ${sigem}`, detail: `${direction}; diferença de revisão identificada.`, sgp, sigem, n2064: left !== null && right !== null };
}

