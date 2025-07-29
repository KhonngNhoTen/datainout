import { PaperSize, Style as CellStyle, Cell, Location } from "exceljs";
import { AttributeType, SheetSection } from "./common-type.js";

export type LayoutSheet<T> = {
  cells: {
    header?: T[];
    footer?: T[];
    table: T[];
  };
  beginTableAt: number;
  endTableAt: number;
  columnIndex: number;
};

export type BaseCellAttribute = {
  section: SheetSection;
  row: number;
  col: number;
  address?: string;
  key: string;
};
export type ImportCell = {
  type: AttributeType;
  required?: boolean;
  setValue?: (attributeValue?: any, row?: Record<string, any>) => any;
  validate?: (val: any) => { isValid: boolean; message?: string } | Error;
} & BaseCellAttribute;

export enum ExportCellTypeEnum {
  VARIABLE = "VARIABLE",
  FORMULA = "FORMULA",
  LABEL = "LABEL",
}

export type ExportCell = {
  value: any;
  type: ExportCellTypeEnum;
  style?: Partial<CellStyle>;
  formatValue?: (value: any) => any;
} & BaseCellAttribute;

export type LayoutExportSheet = {
  merges?: Record<string, { model: Location }>;
  columnWidths?: (number | undefined)[];
  rowHeights?: Record<string, number>;
} & LayoutSheet<ExportCell>;
