import {
  canonicalId, compareNormativeCodes, displayDocumentCode, documentKind, documentTag,
  etDocumentSubtype, identityKey, isEtDocument, norm, ntNeutralKey, ntPrefixVariant,
  parseDocumentReference, parseNormativeDocument, revisionId, revisionInfo, searchKeys, text,
} from "./normative-parser.js";

export {
  canonicalId, displayDocumentCode, documentKind, documentTag, etDocumentSubtype,
  identityKey, isEtDocument, norm, ntNeutralKey, ntPrefixVariant,
  parseDocumentReference, revisionId, searchKeys, text,
};

export const SOURCE_LABELS = { sgp: "SGP", sigem: "Consulta Geral do SIGEM", planned: "Documentos previstos" };

export function allocationState(value) {
  const raw = text(value);
  const key = norm(raw).replace(/[.!;:]+$/g, "").trim();
  if (!key) return { kind: "empty", label: "Sem informação", raw };
  if (/^(?:NAO|N)\s*[-/]?\s*ALOCAD[OA]$/.test(key) || ["NAO", "RECUSADO", "REJEITADO"].includes(key)) return { kind: "not_allocated", label: "Não alocado", raw };
  if (/^ALOCAD[OA](?:\s|$)/.test(key) || ["SIM", "ACEITO", "APROVADO", "CONFIRMADO", "OK"].includes(key)) return { kind: "allocated", label: "Alocado", raw };
  return { kind: "unknown", label: "Revisar valor", raw };
}

function recordScore(record) {
  return [record.title, record.revision, record.status, record.allocationStatus, record.allocationNumber].filter((value) => text(value)).length;
}

function bestRecord(records) {
  let best = null;
  let score = -1;
  for (const record of records) {
    const current = recordScore(record);
    if (current > score || (current === score && (Number(record.rowNumber) || 0) > (Number(best?.rowNumber) || 0))) {
      best = record;
      score = current;
    }
  }
  return best;
}

function unique(values) { return [...new Set(values.filter(Boolean))]; }

function uniqueCodes(records) {
  const codes = new Map();
  for (const record of records) {
    const key = canonicalId(record.code);
    if (key && !codes.has(key)) codes.set(key, displayDocumentCode(key));
  }
  return [...codes.values()];
}

function sourceSummary(records) {
  const codes = uniqueCodes(records);
  const revisions = unique(records.map((record) => revisionId(record.revision)));
  return {
    present: records.length > 0, count: records.length, codes, revisions,
    ambiguous: codes.length > 1 || revisions.length > 1 || records.some((record) => record.revisionConflict),
    record: bestRecord(records), records,
  };
}

function compareField(label, leftValue, rightValue) {
  const sgp = text(leftValue);
  const sigem = text(rightValue);
  if (!sgp && !sigem) return null;
  if (!sgp || !sigem) return { label, kind: "missing", sgp, sigem };
  return norm(sgp) === norm(sigem) ? { label, kind: "equal", sgp, sigem } : { label, kind: "different", sgp, sigem };
}

function sourceFieldComparison(label, sgp, sigem, planned, field) {
  const values = {
    sgp: field === "code" ? sgp.codes.join(" | ") : field === "revision" ? sgp.revisions.join(" | ") : text(sgp.record?.[field]),
    sigem: field === "code" ? sigem.codes.join(" | ") : field === "revision" ? sigem.revisions.join(" | ") : text(sigem.record?.[field]),
    planned: field === "code" ? planned.codes.join(" | ") : field === "revision" ? planned.revisions.join(" | ") : text(planned.record?.[field]),
  };
  if (!sgp.present || !sigem.present) return { label, kind: "not_comparable", values, detail: "Sem as duas fontes para comparar" };
  if (!values.sgp || !values.sigem) return { label, kind: "missing", values, detail: `SGP “${values.sgp || "vazio"}” × SIGEM “${values.sigem || "vazio"}”` };
  const same = field === "code"
    ? sgp.codes.length === sigem.codes.length && sgp.codes.every((code) => sigem.codes.some((other) => canonicalId(code) === canonicalId(other)))
    : norm(values.sgp) === norm(values.sigem);
  return { label, kind: same ? "equal" : "different", values, detail: same ? "Sem diferença" : `SGP “${values.sgp}” × SIGEM “${values.sigem}”` };
}

function resolveAllocation(planned) {
  return planned.present
    ? { kind: "allocated", label: "Alocado", reason: "", evidence: "" }
    : { kind: "not_allocated", label: "Não alocado", reason: "", evidence: "" };
}

function exactSet(summary) { return new Set(summary.codes.map((code) => canonicalId(code))); }

function ntForms(summary) {
  const forms = new Set();
  for (const code of summary.codes) {
    const parsed = parseNormativeDocument(code);
    if (parsed.kind === "et") forms.add(parsed.nt ? "with" : "without");
  }
  return forms;
}

function postingStatus(sgp, sigem, planned) {
  if (!sigem.present) return { kind: "not_posted", label: "Não postado", detail: planned.present ? "Alocado, mas não localizado no SIGEM." : "Não localizado no SIGEM." };
  if (!sgp.present) return { kind: "sigem_only", label: "Postado no SIGEM; ausente no SGP", detail: planned.present ? "Alocado e localizado somente no SIGEM." : "Localizado somente no SIGEM e não alocado." };

  const sgpCodes = exactSet(sgp);
  const sigemCodes = exactSet(sigem);
  const exact = [...sgpCodes].filter((code) => sigemCodes.has(code));
  const changed = [...sigemCodes].filter((code) => !sgpCodes.has(code));
  const sigemForms = ntForms(sigem);
  const sgpForms = ntForms(sgp);

  if (sigemForms.size > 1 || (exact.length && changed.length)) return { kind: "both_forms", label: "As duas formas foram postadas", detail: "O SIGEM contém mais de uma forma do mesmo TAG." };
  if (sgpCodes.size > 1 && exact.length > 0 && exact.length < sgpCodes.size) return { kind: "partial_forms", label: "Parte das formas foi postada", detail: "O SGP contém mais de uma forma e o SIGEM contém apenas parte delas." };
  if (exact.length && !changed.length) return { kind: "same_code", label: "Postado como no SGP", detail: "Código coincidente entre SGP e SIGEM." };
  const ntChanged = sgpForms.size && sigemForms.size && [...sgpForms].some((form) => !sigemForms.has(form));
  return { kind: "changed_code", label: "Postado com código alterado", detail: ntChanged ? "Mesmo TAG postado em outra forma com/sem nt-." : "Mesmo documento localizado com estrutura de código diferente." };
}

function ntExplanation(sgp, sigem) {
  const allCodes = [...sgp.codes.map((code) => ({ source: "SGP", code })), ...sigem.codes.map((code) => ({ source: "SIGEM", code }))];
  const etCodes = allCodes.filter(({ code }) => isEtDocument(code));
  if (!allCodes.length) return { kind: "not_applicable", label: "Sem base SGP/SIGEM", explanation: "Sem código para verificar nt-." };
  if (!etCodes.length) return { kind: "not_applicable", label: "Não se aplica", explanation: "A forma nt- aplica-se aos documentos ET." };
  const forms = new Set(etCodes.map(({ code }) => parseNormativeDocument(code).nt ? "with" : "without"));
  if (forms.size > 1) return { kind: "alternate", label: "Encontrado nas duas formas", explanation: "Há ocorrências com nt- e sem nt- para o mesmo TAG." };
  if (forms.has("with")) return { kind: "exact", label: "Encontrado com nt-", explanation: "O código foi encontrado com o prefixo nt- em minúsculo na apresentação." };
  return { kind: "exact", label: "Encontrado sem nt-", explanation: "O código foi encontrado sem o prefixo nt-." };
}

function presenceStatus(sgp, sigem, planned) {
  if (sgp.present && sigem.present) return { kind: "both", label: "Presente no SGP e SIGEM" };
  if (sgp.present) return { kind: "sgp_only", label: "Somente no SGP" };
  if (sigem.present) return { kind: "sigem_only", label: "Somente no SIGEM" };
  if (planned.present) return { kind: "missing_both", label: "Ausente no SGP e SIGEM" };
  return { kind: "missing", label: "Não localizado" };
}

function representativeMatch(sgp, sigem) {
  if (!sgp.present || !sigem.present) return { kind: "missing", label: "Sem comparação", detail: "Documento ausente em uma das fontes." };
  let fallback = null;
  for (const left of sgp.codes) {
    for (const right of sigem.codes) {
      const match = compareNormativeCodes(left, right);
      if (match.kind === "exact") return match;
      if (!fallback || fallback.kind === "different") fallback = match;
    }
  }
  return fallback || { kind: "different", label: "Código diferente", detail: "Não foi possível estabelecer equivalência." };
}

export function analyzeDatasets(datasets = {}) {
  const groups = new Map();
  for (const source of ["sgp", "sigem", "planned"]) {
    for (const record of datasets[source] || []) {
      const key = identityKey(record.code);
      if (!key) continue;
      if (!groups.has(key)) groups.set(key, { sgp: [], sigem: [], planned: [] });
      groups.get(key)[source].push(record);
    }
  }

  const results = [];
  for (const [groupKey, group] of groups) {
    const sgp = sourceSummary(group.sgp);
    const sigem = sourceSummary(group.sigem);
    const planned = sourceSummary(group.planned);
    const displayCode = displayDocumentCode(sgp.record?.code || sigem.record?.code || planned.record?.code);
    const presence = presenceStatus(sgp, sigem, planned);
    const allocation = resolveAllocation(planned);
    const posting = postingStatus(sgp, sigem, planned);
    const match = representativeMatch(sgp, sigem);
    const revisions = revisionInfo(sgp.revisions.join(" | "), sigem.revisions.join(" | "));
    const normative = parseNormativeDocument(displayCode);
    const fieldComparisons = {
      code: sourceFieldComparison("Código", sgp, sigem, planned, "code"),
      revision: sourceFieldComparison("Revisão", sgp, sigem, planned, "revision"),
      title: sourceFieldComparison("Título", sgp, sigem, planned, "title"),
      status: sourceFieldComparison("Status", sgp, sigem, planned, "status"),
    };
    const comparisons = [
      compareField("Título", sgp.record?.title, sigem.record?.title),
      compareField("Revisão", sgp.revisions.join(" | "), sigem.revisions.join(" | ")),
      compareField("Status", sgp.record?.status, sigem.record?.status),
    ].filter(Boolean);
    const ambiguous = sgp.ambiguous || sigem.ambiguous;
    const differenceDetail = [match.kind === "exact" ? "Código: sem diferença" : `Código: ${match.detail}`, `Revisão: ${revisions.detail}`].join("; ");
    const needsReview = ambiguous || ["changed_code", "both_forms", "partial_forms"].includes(posting.kind) || revisions.kind === "different" || normative.issues.length > 0;
    results.push({
      id: groupKey, groupKey, displayCode, sgp, sigem, planned, presence, allocation, posting, match,
      differenceDetail, revisionComparison: revisions,
      n2064: { valid: revisions.n2064, detail: revisions.n2064 ? "Revisão compatível com a sequência da N-2064." : "Revisão requer validação pela N-2064." },
      normative, documentTag: documentTag(displayCode), comparisons,
      differences: comparisons.filter((item) => item.label === "Revisão" && item.kind !== "equal"),
      fieldComparisons, documentKind: normative.kind, nt: ntExplanation(sgp, sigem), ambiguous, needsReview,
    });
  }
  results.sort((left, right) => left.documentTag.localeCompare(right.documentTag, "pt-BR", { numeric: true }));
  return { results, metrics: summarizeResults(results) };
}

export function summarizeResults(results = []) {
  return {
    total: results.length,
    inBoth: results.filter((item) => item.presence.kind === "both").length,
    onlySgp: results.filter((item) => item.presence.kind === "sgp_only").length,
    onlySigem: results.filter((item) => item.presence.kind === "sigem_only").length,
    onlyPlanned: results.filter((item) => item.presence.kind === "missing_both").length,
    allocated: results.filter((item) => item.allocation.kind === "allocated").length,
    notAllocated: results.filter((item) => item.allocation.kind === "not_allocated").length,
    missingPlanned: results.filter((item) => !item.planned.present).length,
    review: results.filter((item) => item.needsReview).length,
    ntAlternates: results.filter((item) => item.nt.kind === "alternate").length,
    posted: results.filter((item) => item.sigem.present).length,
    notPosted: results.filter((item) => !item.sigem.present).length,
    codeChanged: results.filter((item) => ["changed_code", "both_forms", "partial_forms"].includes(item.posting.kind)).length,
    bothForms: results.filter((item) => item.posting.kind === "both_forms").length,
    revisionChanged: results.filter((item) => item.revisionComparison.kind === "different").length,
    eapChanged: results.filter((item) => item.match.kind === "eap_transition").length,
  };
}

export function filterResults(results, filter, query) {
  const search = norm(query);
  return results.filter((item) => {
    const matchesFilter = filter === "all"
      || (filter === "not_posted" && !item.sigem.present)
      || (filter === "changed" && ["changed_code", "both_forms", "partial_forms"].includes(item.posting.kind))
      || (filter === "not_allocated" && item.allocation.kind === "not_allocated")
      || (filter === "sigem_only" && item.presence.kind === "sigem_only")
      || (filter === "review" && item.needsReview);
    if (!matchesFilter) return false;
    if (!search) return true;
    return norm([item.documentTag, item.displayCode, item.sgp.codes.join(" "), item.sigem.codes.join(" "), item.presence.label, item.allocation.label, item.posting.label, item.differenceDetail].join(" ")).includes(search);
  });
}
