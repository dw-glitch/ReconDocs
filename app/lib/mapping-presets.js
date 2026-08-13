import { normalizedHeader } from "./sheet-profile.js";

/**
 * Assinatura estável do layout de uma planilha: os cabeçalhos normalizados,
 * na ordem em que aparecem. Duas planilhas com o mesmo layout de colunas —
 * mesmo se o arquivo tiver nome ou data diferentes, como a atualização
 * diária de Documentos Previstos — produzem a mesma assinatura.
 */
export function headerSignature(headers = []) {
  return headers.map((header) => normalizedHeader(header)).join("|");
}

const STORAGE_KEY = "recondocs.cruzamento.presets.v1";

function readAll(storage) {
  try {
    const raw = storage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

/**
 * Lê o mapeamento salvo para o layout de cabeçalhos indicado, se houver.
 * @param storage objeto compatível com `localStorage` (getItem/setItem).
 */
export function readMappingPreset(storage, headers) {
  if (!storage) return null;
  const signature = headerSignature(headers);
  if (!signature) return null;
  const preset = readAll(storage)[signature];
  return preset && typeof preset === "object" ? preset : null;
}

/**
 * Salva o mapeamento de colunas (documento, status, data e extras) para o
 * layout de cabeçalhos indicado, para reaplicar automaticamente da próxima
 * vez que uma planilha com o mesmo layout for carregada.
 */
export function writeMappingPreset(storage, headers, mapping) {
  if (!storage) return;
  const signature = headerSignature(headers);
  if (!signature) return;
  try {
    const all = readAll(storage);
    all[signature] = {
      document: mapping.document,
      status: mapping.status,
      date: mapping.date,
      extras: [...(mapping.extras || [])].sort((left, right) => left - right),
    };
    storage.setItem(STORAGE_KEY, JSON.stringify(all));
  } catch {
    // Armazenamento indisponível ou cheio: o preset simplesmente não é salvo.
  }
}
