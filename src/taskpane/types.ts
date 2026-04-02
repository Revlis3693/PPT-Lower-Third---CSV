export type CsvRow = Record<string, string>;

export type ShapeMapping = {
  id: string;
  templateSlideId: string;
  shapeId: string;
  shapeName?: string;
  columnName: string;
  label?: string;
};

export type StatusMessage = {
  kind: "info" | "warning" | "error" | "success";
  message: string;
};

