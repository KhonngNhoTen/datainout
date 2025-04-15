import { PaperSize, Style as CellStyle, Cell, Worksheet, Location } from "exceljs";
import { BaseAttribute, SheetSection } from "./common-type.js";

/** Excel report template */
export type CellReportOptions = {
  address: string;
  fullAddress: Cell["fullAddress"];
  section?: SheetSection;
  value: {
    hardValue?: any;
    fieldName?: string;
  };
  style: Partial<CellStyle>;
  isVariable: boolean;
};

export type SheetReportOptions = {
  pageSize?: PaperSize;
  merges?: Record<string, { model: Location }>;
  beginTableAt: number;
  endTableAt?: number;
  cells: CellReportOptions[];
  columnWidths?: (number | undefined)[];
  rowHeights: Record<string, number>;
};

export type TableReportOptions = {
  sheets: SheetReportOptions[];
  name: string;
};
