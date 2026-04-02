/**
 * Max characters per cell before showing a warning in the preview table.
 * Add or override keys to match your CSV header names. Falls back to `default`.
 */
export const DEFAULT_CHAR_THRESHOLD = 50;

export const COLUMN_CHAR_THRESHOLDS: Record<string, number> = {
  default: 50,
  Name: 48,
  Title: 64,
  Company: 80
};

export function getCharThresholdForColumn(columnName: string): number {
  return COLUMN_CHAR_THRESHOLDS[columnName] ?? COLUMN_CHAR_THRESHOLDS.default ?? DEFAULT_CHAR_THRESHOLD;
}
