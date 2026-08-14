import test from "node:test";
import assert from "node:assert/strict";
import {
  allocationState,
  analyzeDatasets,
  canonicalId,
  documentKind,
  documentTag,
  etDocumentSubtype,
  identityKey,
  ntPrefixVariant,
  parseDocumentReference,
  searchKeys,
} from "../app/lib/analysis-engine.js";

function row(code, extra = {}) {
  return {
    code,
    fileName: "base.xlsx",
    sheetName: "Dados",
    rowNumber: 2,
    allocationStatusHeader: "Confirmação de DOCUMENTOS PREVISTOS",
    ...extra,
  };
}

test("normaliza extensão e espaços sem destruir o código", () => {
  assert.equal(canonicalId(" C1O_RNEST_U32_3.1.1.1_INS_RIR_AB-001.pdf "), "C1O_RNEST_U32_3.1.1.1_INS_RIR_AB-001");
});

test("separa revisão anexada ao documento sem contaminar o código", () => {
  const parsed = parseDocumentReference("ET-5290.00-22000-912-1LV-001 - Rev. P(1).pdf");
  assert.equal(parsed.code, "ET-5290.00-22000-912-1LV-001");
  assert.equal(parsed.revision, "P");
  assert.equal(canonicalId("ET-5290.00-22000-912-1LV-001 - Rev. P(1).pdf"), "ET-5290.00-22000-912-1LV-001");
  assert.deepEqual(
    parseDocumentReference("5900.0018047.05.2-ABC-CV-GER-0001-P").code,
    "5900.0018047.05.2-ABC-CV-GER-0001",
  );
});

test("prioriza a coluna de revisão e sinaliza conflito com a referência", () => {
  const parsed = parseDocumentReference("ET-5290.00-22000-912-1LV-001 Rev. P", "Q");
  assert.equal(parsed.revision, "Q");
  assert.equal(parsed.revisionConflict, true);
});

test("classifica ET, CV e N-1710 e os subtipos ET", () => {
  const rir = "C1O_RNEST_U32_3.1.1.1_INS_RIR_SPE-AST-320019";
  const cm = "C1O_RNEST_U32_3.1.1.1_MEC_CCM_EQ-001";
  assert.equal(documentKind(rir), "et");
  assert.equal(etDocumentSubtype(rir), "rir");
  assert.equal(etDocumentSubtype(cm), "cm");
  assert.equal(documentKind("5900.0018047.05.2-ABC-CV-GER-0001"), "cv");
  assert.equal(documentKind("ET-5290.00-22000-912-1LV-001"), "n1710");
});

test("documento ET recebe busca nas duas formas com/sem NT-", () => {
  const base = "C1O_RNEST_U32_3.1.1.1_INS_RIR_SPE-AST-320019";
  const alternate = "C1O_RNEST_U32_3.1.1.1_INS_RIR_nt-SPE-AST-320019";
  assert.equal(ntPrefixVariant(base), alternate);
  assert.deepEqual(searchKeys(base), [base, alternate]);
  assert.equal(identityKey(base), identityKey(alternate));
  const standardExample = "XXX_RNEST_U34_3.1.1.1_CVL_RIR_nt-EMT-XE-034-004";
  assert.equal(
    ntPrefixVariant("XXX_RNEST_U34_3.1.1.1_CVL_RIR_EMT-XE-034-004"),
    standardExample,
  );
});

test("usa o tag técnico como vínculo sem misturar código e revisão", () => {
  assert.equal(
    documentTag("C1O_RNEST_U32_3.1.1.1_CVL_RIR_nt-EMT-F-32501-PLA-3.200"),
    "EMT-F-32501-PLA-3.200",
  );
  assert.equal(documentTag("PR-5290.00-22313-970-C1O-005"), "PR-5290.00-22313-970-C1O-005");
});

test("N-1710 não recebe prefixo NT-", () => {
  const code = "ET-5290.00-22000-912-1LV-001";
  assert.deepEqual(searchKeys(code), [code]);
  assert.equal(ntPrefixVariant(code), "");
});

test("usa Documentos Previstos para confirmar a alocação", () => {
  const code = "C1O_RNEST_U32_3.1.1.1_INS_RIR_SPE-AST-320019";
  const analyzed = analyzeDatasets({
    sgp: [row(code, { title: "RELATÓRIO A", revision: "A" })],
    sigem: [row(code, { title: "RELATÓRIO A", revision: "A" })],
    planned: [row(code, { allocationStatus: "ALOCADO", allocationNumber: "C1O-ALOC-CM-0001-2026" })],
  });
  assert.equal(analyzed.metrics.total, 1);
  assert.equal(analyzed.results[0].presence.kind, "both");
  assert.equal(analyzed.results[0].allocation.kind, "allocated");
  assert.equal(analyzed.results[0].fieldComparisons.revision.kind, "equal");
});

test("localiza ET quando SGP e SIGEM usam formas diferentes", () => {
  const withoutNt = "C1O_RNEST_U32_3.1.1.1_INS_RIR_SPE-AST-320019";
  const withNt = "C1O_RNEST_U32_3.1.1.1_INS_RIR_nt-SPE-AST-320019";
  const analyzed = analyzeDatasets({
    sgp: [row(withoutNt)],
    sigem: [row(withNt)],
    planned: [row(withoutNt, { allocationStatus: "NÃO ALOCADO" })],
  });
  assert.equal(analyzed.metrics.total, 1);
  assert.equal(analyzed.results[0].nt.kind, "alternate");
  assert.equal(analyzed.results[0].nt.label, "Encontrado nas duas formas");
  assert.equal(analyzed.results[0].displayCode, withoutNt);
  assert.equal(analyzed.results[0].allocation.kind, "allocated");
});

test("apresenta nt- sempre em minúsculo mesmo quando a fonte usa maiúsculas", () => {
  const uppercase = "C1O_RNEST_U32_3.1.1.1_INS_RIR_NT-SPE-AST-320019";
  const analyzed = analyzeDatasets({ sgp: [row(uppercase)], sigem: [], planned: [] });
  assert.equal(analyzed.results[0].displayCode, "C1O_RNEST_U32_3.1.1.1_INS_RIR_nt-SPE-AST-320019");
  assert.equal(analyzed.results[0].sgp.codes[0], "C1O_RNEST_U32_3.1.1.1_INS_RIR_nt-SPE-AST-320019");
  assert.equal(analyzed.results[0].nt.label, "Encontrado com nt-");
});

test("informa explicitamente quando encontra somente sem NT-", () => {
  const code = "C1O_RNEST_U32_3.1.1.1_INS_RIR_SPE-AST-320019";
  const analyzed = analyzeDatasets({ sgp: [row(code)], sigem: [row(code)], planned: [] });
  assert.equal(analyzed.results[0].nt.label, "Encontrado sem nt-");
});

test("presença nos previstos basta para informar alocado", () => {
  assert.equal(allocationState("NÃO ALOCADO").kind, "not_allocated");
  const code = "ET-5290.00-22000-912-1LV-001";
  const analyzed = analyzeDatasets({
    sgp: [row(code)],
    sigem: [row(code)],
    planned: [row(code, { allocationStatus: "NÃO ALOCADO", allocationNumber: "C1O-ALOC-CM-0099-2026" })],
  });
  assert.equal(analyzed.results[0].allocation.kind, "allocated");
  assert.equal(analyzed.results[0].allocation.label, "Alocado");
  assert.equal(analyzed.results[0].allocation.reason, "");
});

test("ausência nos previstos informa não alocado", () => {
  const code = "ET-5290.00-22000-912-1LV-001";
  const analyzed = analyzeDatasets({ sgp: [row(code)], sigem: [row(code)], planned: [] });
  assert.equal(analyzed.results[0].allocation.kind, "not_allocated");
  assert.equal(analyzed.results[0].allocation.label, "Não alocado");
});

test("liga códigos diferentes pelo mesmo tag e informa como foi alocado e postado", () => {
  const withNt = "C1O_RNEST_U32_3.1.1.1_CVL_RIR_nt-EMT-F-32501-PLA-3.200";
  const withoutNt = "C1O_RNEST_U32_3.1.1.1_CVL_RIR_EMT-F-32501-PLA-3.200";
  const analyzed = analyzeDatasets({
    sgp: [row(withNt, { revision: "A" })],
    sigem: [row(withoutNt, { revision: "0" })],
    planned: [row(withNt), row(withoutNt)],
  });
  assert.equal(analyzed.metrics.total, 1);
  assert.equal(analyzed.results[0].documentTag, "EMT-F-32501-PLA-3.200");
  assert.equal(analyzed.results[0].posting.kind, "changed_code");
  assert.doesNotMatch(analyzed.results[0].posting.detail, /Documentos Previstos|Alocado e postado/);
  assert.equal(analyzed.results[0].fieldComparisons.revision.kind, "different");
});

test("informa quando as duas formas do mesmo tag foram postadas", () => {
  const withNt = "C1O_RNEST_U32_3.1.1.1_CVL_RIR_nt-B-32006A";
  const withoutNt = "C1O_RNEST_U32_3.1.1.1_CVL_RIR_B-32006A";
  const analyzed = analyzeDatasets({
    sgp: [row(withNt)],
    sigem: [row(withNt), row(withoutNt)],
    planned: [row(withNt), row(withoutNt)],
  });
  assert.equal(analyzed.results[0].posting.kind, "both_forms");
  assert.equal(analyzed.results[0].posting.label, "As duas formas foram postadas");
});

test("reconhece código N-1710 encapsulado no nome do checklist do SGP", () => {
  const sgpCode = "C1O_RNEST_U32_3.1.1.1_ELE_LBR-G18-M1.1-098-104480_PR-5290.00-22313-970-C1O-005";
  const sigemCode = "PR-5290.00-22313-970-C1O-005";
  const analyzed = analyzeDatasets({ sgp: [row(sgpCode)], sigem: [row(sigemCode)], planned: [row(sigemCode)] });
  assert.equal(analyzed.metrics.total, 1);
  assert.equal(analyzed.results[0].posting.kind, "changed_code");
  assert.equal(analyzed.results[0].documentTag, sigemCode);
});

test("cruza mais de 20 mil documentos por índices em tempo linear", () => {
  const count = 22000;
  const sgp = [];
  const sigem = [];
  const planned = [];
  for (let index = 0; index < count; index += 1) {
    const code = `C1O_RNEST_U32_3.1.1.1_INS_RIR_ITEM-${String(index).padStart(6, "0")}`;
    sgp.push(row(code));
    sigem.push(row(code));
    planned.push(row(code, { allocationStatus: index % 2 ? "ALOCADO" : "NÃO ALOCADO" }));
  }
  const start = performance.now();
  const analyzed = analyzeDatasets({ sgp, sigem, planned });
  const elapsed = performance.now() - start;
  assert.equal(analyzed.metrics.total, count);
  assert.equal(analyzed.metrics.inBoth, count);
  assert.ok(elapsed < 5000, `análise levou ${elapsed.toFixed(0)} ms`);
});
