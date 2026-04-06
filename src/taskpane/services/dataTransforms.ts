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

/** True if at least one column has non-whitespace content. */
export function rowHasAnyCellData(row: CsvRow, headers: string[]): boolean {
  return headers.some((h) => {
    const v = row[h];
    return v != null && String(v).trim().length > 0;
  });
}

/** Drop rows where every column is empty or whitespace (after values are as stored in row). */
export function dropRowsWithNoCellData(rows: CsvRow[], headers: string[]): CsvRow[] {
  return rows.filter((row) => rowHasAnyCellData(row, headers));
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
