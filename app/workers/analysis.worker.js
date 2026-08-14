import { analyzeDatasets } from "../lib/analysis-engine.js";

self.onmessage = (event) => {
  try {
    self.postMessage({ ok: true, output: analyzeDatasets(event.data) });
  } catch (error) {
    self.postMessage({ ok: false, error: error instanceof Error ? error.message : "Falha na análise." });
  }
};

