import Papa from "papaparse";
import type { CsvRow } from "../types";
import { dropRowsWithNoCellData } from "./dataTransforms";

export type ParsedCsv = {
  headers: string[];
  rows: CsvRow[];
};

export function parseCsv(file: File): Promise<ParsedCsv> {
  return new Promise((resolve, reject) => {
    Papa.parse<CsvRow>(file, {
      header: true,
      // "greedy" also skips lines that are only commas / delimiters (no cell text).
      skipEmptyLines: "greedy",
      transformHeader: (h) => h.trim(),
      transform: (v) => (typeof v === "string" ? v.trim() : v),
      complete: (results) => {
        if (results.errors?.length) {
          reject(new Error(results.errors.map((e) => e.message).join("\n")));
          return;
        }
        const metaFields = results.meta?.fields ?? [];
        const headers = metaFields.filter((h) => h && h.trim().length > 0);

        const rows = (results.data ?? []).map((r) => {
          const out: CsvRow = {};
          for (const h of headers) out[h] = (r as any)?.[h] ?? "";
          return out;
        });

        resolve({
          headers,
          rows: dropRowsWithNoCellData(rows, headers)
        });
      },
      error: (err) => reject(err)
    });
  });
}

