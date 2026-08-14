import assert from "node:assert/strict";
import test from "node:test";
import {
  compareNormativeCodes,
  documentKind,
  identityKey,
  parseNormativeDocument,
  revisionInfo,
} from "../app/lib/normative-parser.js";

test("separa os sete grupos ET e mantém o TAG estável", () => {
  const left = "C1O_RNEST_U32_7.1.1.1_INS_GRACIM_nt-SPE-AST-320019";
  const right = "C1O_RNEST_U32_7.8.1.1_INS_RRIMTI_SPE-AST-320019";
  const parsed = parseNormativeDocument(left);
  assert.equal(parsed.kind, "et");
  assert.equal(parsed.eap, "7.1.1.1");
  assert.equal(parsed.eapCurrent, "7.8.1.1");
  assert.equal(parsed.tag, "SPE-AST-320019");
  assert.equal(identityKey(left), identityKey(right));
  assert.equal(compareNormativeCodes(left, right).kind, "eap_transition");
});

test("não classifica texto desconhecido como N-1710", () => {
  assert.equal(documentKind("DOCUMENTO-LIVRE-123"), "unknown");
  assert.equal(documentKind("PR-5290.00-22313-970-C1O-005"), "n1710");
});

test("aplica a sequência de revisão da N-2064 sem I e O", () => {
  assert.equal(revisionInfo("H", "J").kind, "different");
  assert.equal(revisionInfo("H", "J").n2064, true);
  assert.equal(revisionInfo("I", "J").n2064, false);
});

