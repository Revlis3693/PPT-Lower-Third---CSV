import type { CsvRow } from "../types";

/** Trim ends and collapse internal whitespace runs to a single space. */
export function cleanRow(row: CsvRow, headers: string[]): CsvRow {
  const out: CsvRow = {};
  for (const h of headers) {
    let v = row[h] ?? "";
    v = v.trim();
    v = v.replace(/\s+/g, " ");
    out[h] = v;
  }
  return out;
}

/** Remove duplicate rows where all column values match (order-preserving). */
export function dedupeRows(rows: CsvRow[], headers: string[]): CsvRow[] {
  const seen = new Set<string>();
  const out: CsvRow[] = [];
  for (const r of rows) {
    const key = headers.map((h) => r[h] ?? "").join("\u0001");
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(r);
  }
  return out;
}
