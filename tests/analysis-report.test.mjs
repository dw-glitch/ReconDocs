import assert from "node:assert/strict";
import test from "node:test";
import { analyzeDatasets } from "../app/lib/analysis-engine.js";
import { ANALYSIS_REPORT_COLUMNS, analysisReportRows, buildAnalysisWorkbook } from "../app/lib/analysis-report.js";

const code = "C1O_RNEST_U32_3.1.1.1_INS_RIR_nt-SPE-AST-320019";
const record = (value, revision = "") => ({ code: value, revision, fileName: "x.xlsx", sheetName: "Dados", rowNumber: 2 });

test("relatório expõe somente as nove colunas essenciais", async () => {
  const output = analyzeDatasets({ sgp: [record(code, "A")], sigem: [record(code.replace("nt-", ""), "0")], planned: [record(code)] });
  assert.deepEqual(ANALYSIS_REPORT_COLUMNS.map((column) => column.header), [
    "TAG DO DOCUMENTO", "CÓDIGO NO SGP", "REVISÃO SGP", "CÓDIGO POSTADO NO SIGEM",
    "REVISÃO SIGEM", "ALOCAÇÃO", "SITUAÇÃO DA POSTAGEM", "DIFERENÇA IDENTIFICADA", "FORMA nt-",
  ]);
  const [row] = analysisReportRows(output.results);
  assert.equal(row.alocacao, "Alocado");
  assert.ok(!Object.keys(row).some((key) => /previsto|codigoAlocado/i.test(key)));
  const workbook = await buildAnalysisWorkbook(output.results, output.metrics);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Resumo", "Conferência"]);
  const buffer = await workbook.xlsx.writeBuffer();
  assert.ok(buffer.byteLength > 5000);
});

