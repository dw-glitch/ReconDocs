import test from "node:test";
import assert from "node:assert/strict";
import { headerSignature, readMappingPreset, writeMappingPreset } from "../app/lib/mapping-presets.js";

function fakeStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, value),
  };
}

test("a assinatura ignora acentos, caixa e espaçamento, mas respeita a ordem", () => {
  assert.equal(
    headerSignature(["Número do Documento", "Status", "Data de Postagem"]),
    headerSignature(["NÚMERO DO DOCUMENTO", "status", "  Data   de   Postagem  "]),
  );
  assert.notEqual(
    headerSignature(["Documento", "Status"]),
    headerSignature(["Status", "Documento"]),
  );
});

test("salva e recupera o mapeamento pelo layout de cabeçalhos", () => {
  const storage = fakeStorage();
  const headers = ["Número do Documento", "Status", "Data de Postagem", "Disciplina"];
  assert.equal(readMappingPreset(storage, headers), null);

  writeMappingPreset(storage, headers, { document: 0, status: 1, date: 2, extras: [3] });
  const preset = readMappingPreset(storage, headers);
  assert.deepEqual(preset, { document: 0, status: 1, date: 2, extras: [3] });
});

test("planilhas com o mesmo layout de cabeçalho reaproveitam o preset, mesmo com nomes de arquivo diferentes", () => {
  const storage = fakeStorage();
  const headersToday = ["Código do Documento", "Responsável"];
  const headersTomorrow = ["Código do Documento", "Responsável"];
  writeMappingPreset(storage, headersToday, { document: 0, status: -1, date: -1, extras: [1] });
  assert.deepEqual(readMappingPreset(storage, headersTomorrow), { document: 0, status: -1, date: -1, extras: [1] });
});

test("um layout de cabeçalhos diferente não reaproveita o preset de outro", () => {
  const storage = fakeStorage();
  writeMappingPreset(storage, ["Documento", "Status"], { document: 0, status: 1, date: -1, extras: [] });
  assert.equal(readMappingPreset(storage, ["Documento", "Status", "Data"]), null);
});

test("sem storage disponível as funções não lançam erro", () => {
  assert.equal(readMappingPreset(null, ["Documento"]), null);
  assert.doesNotThrow(() => writeMappingPreset(null, ["Documento"], { document: 0, status: -1, date: -1, extras: [] }));
});

test("storage que lança ao ler ou escrever não derruba a aplicação", () => {
  const broken = {
    getItem: () => { throw new Error("bloqueado"); },
    setItem: () => { throw new Error("cheio"); },
  };
  assert.equal(readMappingPreset(broken, ["Documento"]), null);
  assert.doesNotThrow(() => writeMappingPreset(broken, ["Documento"], { document: 0, status: -1, date: -1, extras: [] }));
});
