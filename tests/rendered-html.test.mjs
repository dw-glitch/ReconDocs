import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function rendered(route) {
  const path = route === "/" ? "../.next/server/app/index.html" : `../.next/server/app${route}.html`;
  return readFile(new URL(path, import.meta.url), "utf8");
}

test("renderiza a marca e as três entradas da conferência", async () => {
  const html = await rendered("/");
  assert.match(html, /RECON/);
  assert.match(html, /DOCS/);
  assert.match(html, /Lista do SGP/);
  assert.match(html, /Consulta Geral do SIGEM/);
  assert.match(html, /Documentos previstos/);
});

test("renderiza o módulo limpo de cruzamento", async () => {
  const html = await rendered("/cruzamento");
  assert.match(html, /Cruzamento de planilhas/);
  assert.doesNotMatch(html, /Carregue quantas planilhas quiser/);
});

