export type CellType = "number" | "string" | "boolean" | "object" | "date" | "virtual";
export type SheetSection = "header" | "table" | "footer";

export type CellDescription = {
  type: CellType;
  fieldName: string;
  section: SheetSection;
  address?: string;
  fullAddress: {
    sheetName: string;
    address: string;
    row: number;
    col: number;
  };
  setValue?: (rawCellvalue: string | undefined, row: {}) => any;
  validate?: (val: any) => boolean;
};

export type SheetDesciptionOptions = {
  startTable?: number;
  endTable?: number;
  content: Array<CellDescription>;
  name?: string;
  keyIndex?: number;
};
export type TemplateExcelImportOptions = {
  sheets: SheetDesciptionOptions[];
};

export type FilterImportHandler = {
  sheetIndex: number;
  sheetName?: string;
  section: SheetSection;
};
