/**
 * Catálogo normativo mínimo usado pelo motor do ReconDocs.
 *
 * Fontes funcionais incorporadas ao produto:
 * - ET-5290.00-22000-912-1LV-001, revisões N e P;
 * - Petrobras N-1710;
 * - Petrobras N-2064;
 * - Caminho das Pastas / critério EAP da UHDT-D.
 *
 * O catálogo não tenta substituir as normas. Ele registra apenas regras
 * determinísticas necessárias para reconhecer, vincular e explicar códigos.
 */

export const ET_REPORT_SUCCESSORS = new Map([
  ["GRACIM", "RRIMTI"],
  ["GRACIMR", "RRIMTIR"],
]);

export const ET_REPORT_CODES = new Set([
  "APR", "ARQ", "ASB", "CAL", "CCM", "C&M", "CE", "CL", "CR", "DB",
  "DE", "DF", "DG", "DI", "DL", "DR", "DS", "DT", "EC", "FD", "FL",
  "GRACIM", "GRACIMR", "IM", "IS", "LA", "LD", "LI", "LO", "MA", "MC",
  "MD", "MO", "PR", "PT", "RIR", "RL", "RM", "RRIMTI", "RRIMTIR",
]);

export const N1710_CATEGORIES = new Set([
  "CE", "CR", "DB", "DE", "EC", "ET", "FD", "IM", "IS", "LA", "LD",
  "LI", "LO", "MA", "MC", "MD", "MO", "PR", "PT", "RL", "RM",
]);

// A árvore antiga (7.x) aparece em entregas anteriores. O caminho oficial
// fornecido para a UHDT-D posiciona o mesmo ramo sob 7.8.x.
const LEGACY_EAP_BRANCHES = [
  ["7.1", "7.8.1", 7],
  ["7.2", "7.8.2", 1],
  ["7.3", "7.8.3", 5],
  ["7.4", "7.8.4", 2],
  ["7.5", "7.8.5", 7],
  ["7.6", "7.8.6", 4],
  ["7.7", "7.8.7", 1],
];

export const EAP_LEGACY_TO_CURRENT = new Map(
  LEGACY_EAP_BRANCHES.flatMap(([legacy, current, count]) =>
    Array.from({ length: count }, (_, index) => {
      const item = index + 1;
      return [
        [`${legacy}.${item}`, `${current}.${item}`],
        [`${legacy}.${item}.1`, `${current}.${item}`],
      ];
    }).flat(),
  ),
);

export const EAP_CURRENT_TO_LEGACY = new Map(
  [...EAP_LEGACY_TO_CURRENT].map(([legacy, current]) => [current, legacy]),
);

export const N2064_REVISION_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ";
