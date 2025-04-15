export type TableData = {
  header?: {};
  footer?: {};
  table?: any[];
};

export type PageData = Record<string, any>;

export type DataInoutInput = TableData | PageData;

export type AttributeType = "number" | "string" | "boolean" | "object" | "date" | "virtual";
export type BaseAttribute = {
  type: AttributeType;
  keyName: string;
};

export type SheetSection = "header" | "table" | "footer";
