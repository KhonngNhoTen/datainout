export type TableData = {
  header?: any;
  footer?: any;
  table?: any[];
};

export type PageData = Record<string, any>;

export type DataInoutInput = TableData | PageData;

export type AttributeType = "number" | "string" | "boolean" | "object" | "date" | "virtual";
export type BaseAttribute = {
  keyName: string;
  index?: number;
  section: SheetSection;
};

export type SheetSection = "header" | "table" | "footer";

export type SheetExcelOption<T extends BaseAttribute> = {
  cells: T[];
  sheetIndex: number;
  sheetName: string;
  beginTableAt: number;
  endTableAt?: number;
  keyTableAt: number;
};

export type EventType = {
  finish: () => void;
  begin: (sheetName?: string) => void;
  data: (value: any) => void;
  end: (sheetName?: string) => void;

  /** Handle error. Return false to cancel import, otherhands return true */
  error: (error: Error) => boolean;
};

export type TableExcelOptions<T> = {
  sheets: T[];
  name: string;
};
