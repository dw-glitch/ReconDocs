export const SOURCE_LABELS = {
  sgp: "SGP",
  sigem: "Consulta Geral do SIGEM",
  planned: "Documentos previstos",
};

const ET_PATTERN = /^(?:[^_]+_RNEST_[^_]+_\d+(?:\.\d+){3}_[^_]+_[^_]+_)(.+)$/;
const KNOWN_FILE_EXTENSION = /\.(?:PDF|DWG|DGN|DOCX?|XLSX?|XLSM|CSV|TSV|ZIP)$/i;
const REVISION_SUFFIX = /(?:\s*[-–—]?\s*)(?:REV(?:IS(?:ÃO|AO))?\.?|REVISÃO)\s*[:.\-]?\s*([A-Z0-9]{1,5})(?:\s*\(\d+\))?\s*$/i;
const STANDARD_CODE_WITH_REVISION = /^((?:[A-Z]{1,4}-\d{4}\.\d{2}-\d{5}-\d{3}-[A-Z0-9]{3}-\d{3}|\d{4}\.\d{7}\.\d{2}\.\d{1,2}-[A-Z0-9]{3}-CV-[A-Z0-9]{3}-\d{3,4}))(?:\s*[-_–—]\s*([A-Z]{1,3})(?:\s*\(\d+\))?)?$/i;

export function text(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Intl.DateTimeFormat("pt-BR").format(value);
  }
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
    .replace(/[\s\-–—]+$/g, "")
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

export function isEtDocument(value) {
  return ET_PATTERN.test(canonicalId(value));
}

export function ntNeutralKey(value) {
  const key = canonicalId(value);
  const match = key.match(ET_PATTERN);
  if (!match) return key;
  const prefix = key.slice(0, key.length - match[1].length);
  return `${prefix}${match[1].replace(/^NT-/, "")}`;
}

export function ntPrefixVariant(value) {
  const key = canonicalId(value);
  const match = key.match(ET_PATTERN);
  if (!match) return "";
  const prefix = key.slice(0, key.length - match[1].length);
  return match[1].startsWith("NT-")
    ? `${prefix}${match[1].slice(3)}`
    : `${prefix}nt-${match[1]}`;
}

export function searchKeys(value) {
  const key = canonicalId(value);
  if (!key) return [];
  const displayedKey = displayDocumentCode(key);
  return isEtDocument(key) ? [...new Set([displayedKey, ntPrefixVariant(displayedKey)].filter(Boolean))] : [displayedKey];
}

export function documentTag(value) {
  const key = canonicalId(value);
  const match = key.match(ET_PATTERN);
  if (!match) return key;
  return match[1].replace(/^NT-/, "");
}

export function identityKey(value) {
  const key = canonicalId(value);
  const tag = documentTag(key);
  return tag ? `TAG::${tag}` : "";
}

export function documentKind(value) {
  const key = canonicalId(value);
  if (!key) return "unknown";
  if (/(?:^|[-_])CV(?:[-_]|$)/.test(key)) return "cv";
  if (isEtDocument(key)) return "et";
  return "n1710";
}

export function etDocumentSubtype(value, title = "") {
  const key = canonicalId(value);
  if (!isEtDocument(key)) return "not_et";
  const groups = key.split("_");
  const reportCode = groups[5] || "";
  const searchable = norm(`${reportCode} ${title}`);
  if (/^RIR(?:-|$)/.test(reportCode) || /\bRIR\b|INSPECAO DE RECEBIMENTO/.test(searchable)) return "rir";
  if (/^(?:C&M|CM|CCM)(?:-|$)/.test(reportCode) || /\bC\s*&\s*M\b|\bC E M\b|COMPLETACAO MECANICA|COMISSIONAMENTO/.test(searchable)) return "cm";
  return "other";
}

export function allocationState(value) {
  const raw = text(value);
  const valueKey = norm(raw).replace(/[.!;:]+$/g, "").trim();
  if (!valueKey) return { kind: "empty", label: "Sem informação", raw };

  const negative = new Set([
    "NAO ALOCADO", "NAO", "RECUSADO", "REJEITADO", "ALOCACAO RECUSADA",
    "DOCUMENTO NAO ALOCADO", "NAO ALOCADA", "PENDENTE DE ALOCACAO",
  ]);
  const positive = new Set([
    "ALOCADO", "SIM", "ACEITO", "APROVADO", "CONFIRMADO", "CONCLUIDO",
    "OK", "DOCUMENTO ALOCADO", "ALOCACAO ACEITA", "ACEITE", "ALOCACAO CONFIRMADA",
  ]);

  if (negative.has(valueKey) || /^(NAO|N)\s*[-/]?\s*ALOCAD[OA]$/.test(valueKey)) {
    return { kind: "not_allocated", label: "Não alocado", raw };
  }
  if (positive.has(valueKey) || /^ALOCAD[OA](?:\s|$)/.test(valueKey)) {
    return { kind: "allocated", label: "Alocado", raw };
  }
  return { kind: "unknown", label: "Revisar valor", raw };
}

function bestRecord(records) {
  return [...records].sort((left, right) => {
    const score = (item) => [item.title, item.revision, item.status, item.allocationStatus, item.allocationNumber]
      .filter((value) => text(value)).length;
    return score(right) - score(left) || (Number(right.rowNumber) || 0) - (Number(left.rowNumber) || 0);
  })[0] || null;
}

function uniqueCodes(records) {
  return [...new Map(records.map((record) => [canonicalId(record.code), displayDocumentCode(record.code)])).values()].filter(Boolean);
}

function sourceSummary(records) {
  const codes = uniqueCodes(records);
  const revisions = [...new Set(records.map((record) => revisionId(record.revision)).filter(Boolean))];
  return {
    present: records.length > 0,
    count: records.length,
    codes,
    revisions,
    ambiguous: codes.length > 1 || revisions.length > 1 || records.some((record) => record.revisionConflict),
    record: bestRecord(records),
  };
}

function compareField(label, leftValue, rightValue) {
  const left = text(leftValue);
  const right = text(rightValue);
  if (!left && !right) return null;
  if (!left || !right) return { label, kind: "missing", sgp: left, sigem: right };
  if (norm(left) === norm(right)) return { label, kind: "equal", sgp: left, sigem: right };
  return { label, kind: "different", sgp: left, sigem: right };
}

function compareAcrossSources(label, sgp, sigem, planned, field) {
  const summaries = { sgp, sigem, planned };
  const values = {};
  for (const source of ["sgp", "sigem", "planned"]) {
    const summary = summaries[source];
    values[source] = field === "code"
      ? summary.codes.join(" | ")
      : field === "revision"
        ? summary.revisions.join(" | ")
        : text(summary.record?.[field]);
  }
  const presentValues = ["sgp", "sigem"]
    .filter((source) => summaries[source].present)
    .map((source) => values[source]);
  const normalized = presentValues.map((value) => field === "code" ? norm(value) : norm(value));
  const nonEmpty = normalized.filter(Boolean);
  let kind = "equal";
  if (presentValues.length < 2) kind = "not_comparable";
  else if (nonEmpty.length !== presentValues.length) kind = "missing";
  else if (new Set(normalized).size > 1) kind = "different";
  const sourceLabels = { sgp: "SGP", sigem: "SIGEM", planned: "Previstos" };
  const detail = kind === "equal"
    ? "Sem diferença"
    : kind === "not_comparable"
      ? "Sem outra fonte presente para comparar"
      : ["sgp", "sigem"]
        .filter((source) => summaries[source].present)
        .map((source) => `${sourceLabels[source]} “${values[source] || "vazio"}”`)
        .join(" × ");
  return { label, kind, values, detail };
}

function resolveAllocation(plannedSummary) {
  if (!plannedSummary.present) {
    return {
      kind: "not_allocated",
      label: "Não alocado",
      reason: "",
      evidence: "",
    };
  }
  return {
    kind: "allocated",
    label: "Alocado",
    reason: "",
    evidence: "",
  };
}

function exactCodeSet(summary) {
  return new Set(summary.codes.map((code) => canonicalId(code)));
}

function postingStatus(sgp, sigem, planned) {
  const sgpCodes = exactCodeSet(sgp);
  const sigemCodes = exactCodeSet(sigem);
  const plannedCodes = planned.codes.map((code) => ({ code, key: canonicalId(code) }));
  const exactSgpInSigem = [...sgpCodes].filter((code) => sigemCodes.has(code));
  const changedSigemCodes = sigem.codes.filter((code) => !sgpCodes.has(canonicalId(code)));
  const postedPlannedCodes = plannedCodes.filter(({ key }) => sigemCodes.has(key)).map(({ code }) => code);
  const notPostedPlannedCodes = plannedCodes.filter(({ key }) => !sigemCodes.has(key)).map(({ code }) => code);

  let kind = "not_posted";
  let label = "Não postado";
  if (sigem.present && !sgp.present) {
    kind = "sigem_only";
    label = "Postado no SIGEM; ausente no SGP";
  } else if (sigem.present && exactSgpInSigem.length && changedSigemCodes.length) {
    kind = "both_forms";
    label = "As duas formas foram postadas";
  } else if (sigem.present && exactSgpInSigem.length) {
    kind = "same_code";
    label = "Postado como no SGP";
  } else if (sigem.present) {
    kind = "changed_code";
    label = "Postado com código alterado";
  }

  const details = [];
  if (postedPlannedCodes.length) details.push(`Alocado e postado: ${postedPlannedCodes.join(" | ")}`);
  if (notPostedPlannedCodes.length) details.push(`Alocado e não postado: ${notPostedPlannedCodes.join(" | ")}`);
  if (!planned.present && sigem.present) details.push("Postado no SIGEM, mas não localizado em Documentos Previstos.");
  if (!sigem.present && planned.present) details.push(`Alocado, mas não postado: ${planned.codes.join(" | ")}`);
  if (!sigem.present && !planned.present) details.push("Não localizado no SIGEM nem em Documentos Previstos.");

  return {
    kind,
    label,
    detail: details.join(" "),
    postedPlannedCodes,
    notPostedPlannedCodes,
  };
}

function ntExplanation(sgp, sigem) {
  const summaries = [
    ["SGP", sgp],
    ["SIGEM", sigem],
  ].filter(([, summary]) => summary.present);
  const allCodes = summaries.flatMap(([label, summary]) => summary.codes.map((code) => ({ label, code })));
  if (!allCodes.length) {
    return { kind: "not_applicable", label: "Sem base SGP/SIGEM", explanation: "O código não foi localizado no SGP nem no SIGEM para verificar a forma com/sem nt-." };
  }
  if (!isEtDocument(allCodes[0].code)) {
    return { kind: "not_applicable", label: "Não se aplica", explanation: "A busca com e sem nt- é exclusiva de documentos ET." };
  }
  const forms = new Set(allCodes.map(({ code }) => canonicalId(code).includes("_NT-") ? "com" : "sem"));
  const searched = searchKeys(allCodes[0].code);
  if (forms.size > 1) {
    const withNt = allCodes.filter(({ code }) => canonicalId(code).includes("_NT-")).map(({ label }) => label);
    const withoutNt = allCodes.filter(({ code }) => !canonicalId(code).includes("_NT-")).map(({ label }) => label);
    return {
      kind: "alternate",
      label: "Encontrado nas duas formas",
      explanation: `Com nt-: ${[...new Set(withNt)].join(", ")}. Sem nt-: ${[...new Set(withoutNt)].join(", ")}. Foram pesquisadas as duas formas (${searched.join(" | ")}) e elas representam o mesmo documento ET.`,
    };
  }
  const foundForm = forms.has("com") ? "com" : "sem";
  const sources = [...new Set(allCodes.map(({ label }) => label))].join(", ");
  return {
    kind: "exact",
    label: foundForm === "com" ? "Encontrado com nt-" : "Encontrado sem nt-",
    explanation: `Encontrado ${foundForm} nt- em ${sources}. A pesquisa também verificou a forma alternativa (${searched.join(" | ")}).`,
  };
}

function presenceStatus(sgp, sigem, planned) {
  if (sgp.present && sigem.present) return { kind: "both", label: "Presente no SGP e SIGEM" };
  if (sgp.present) return { kind: "sgp_only", label: "Somente no SGP" };
  if (sigem.present) return { kind: "sigem_only", label: "Somente no SIGEM" };
  if (planned.present) return { kind: "missing_both", label: "Ausente no SGP e SIGEM" };
  return { kind: "missing", label: "Não localizado" };
}

export function analyzeDatasets(datasets) {
  const groups = new Map();
  for (const source of ["sgp", "sigem", "planned"]) {
    for (const record of datasets[source] || []) {
      const groupKey = identityKey(record.code);
      if (!groupKey) continue;
      if (!groups.has(groupKey)) groups.set(groupKey, { sgp: [], sigem: [], planned: [] });
      groups.get(groupKey)[source].push(record);
    }
  }

  const results = [...groups.entries()].map(([groupKey, group]) => {
    const sgp = { ...sourceSummary(group.sgp), records: group.sgp };
    const sigem = { ...sourceSummary(group.sigem), records: group.sigem };
    const planned = { ...sourceSummary(group.planned), records: group.planned };
    const displayCode = displayDocumentCode(sgp.record?.code || sigem.record?.code || planned.record?.code);
    const presence = presenceStatus(sgp, sigem, planned);
    const allocation = resolveAllocation(planned);
    const posting = postingStatus(sgp, sigem, planned);
    const comparisons = [
      compareField("Título", sgp.record?.title, sigem.record?.title),
      compareField("Revisão", sgp.record?.revision, sigem.record?.revision),
      compareField("Status", sgp.record?.status, sigem.record?.status),
    ].filter(Boolean);
    const differences = comparisons.filter((item) => item.label === "Revisão" && item.kind !== "equal");
    const fieldComparisons = {
      code: compareAcrossSources("Código", sgp, sigem, planned, "code"),
      revision: compareAcrossSources("Revisão", sgp, sigem, planned, "revision"),
      title: compareAcrossSources("Título", sgp, sigem, planned, "title"),
      status: compareAcrossSources("Status", sgp, sigem, planned, "status"),
    };
    const nt = ntExplanation(sgp, sigem);
    const ambiguous = sgp.ambiguous || sigem.ambiguous;
    const needsReview = ambiguous
      || posting.kind === "changed_code"
      || posting.kind === "both_forms"
      || fieldComparisons.revision.kind === "different";

    return {
      id: groupKey,
      displayCode,
      groupKey,
      sgp,
      sigem,
      planned,
      presence,
      allocation,
      posting,
      documentTag: documentTag(displayCode),
      comparisons,
      differences,
      fieldComparisons,
      documentKind: documentKind(displayCode),
      nt,
      ambiguous,
      needsReview,
    };
  }).sort((left, right) => left.displayCode.localeCompare(right.displayCode, "pt-BR", { numeric: true }));

  const metrics = summarizeResults(results);

  return { results, metrics };
}

export function summarizeResults(results) {
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
    codeChanged: results.filter((item) => ["changed_code", "both_forms"].includes(item.posting.kind)).length,
    bothForms: results.filter((item) => item.posting.kind === "both_forms").length,
  };
}

export function filterResults(results, filter, query) {
  const search = norm(query);
  return results.filter((item) => {
    const matchesFilter = filter === "all"
      || (filter === "not_posted" && !item.sigem.present)
      || (filter === "changed" && ["changed_code", "both_forms"].includes(item.posting.kind))
      || (filter === "not_allocated" && item.allocation.kind === "not_allocated")
      || (filter === "sigem_only" && item.presence.kind === "sigem_only")
      || (filter === "review" && (item.needsReview || !["allocated", "not_allocated"].includes(item.allocation.kind)));
    if (!matchesFilter) return false;
    if (!search) return true;
    return norm([
      item.documentTag,
      item.displayCode,
      item.sgp.codes.join(" "),
      item.sigem.codes.join(" "),
      item.planned.codes.join(" "),
      item.presence.label,
      item.allocation.label,
      item.posting.label,
      item.posting.detail,
    ].join(" ")).includes(search);
  });
}
