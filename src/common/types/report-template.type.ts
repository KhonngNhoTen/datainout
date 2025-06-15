import { PaperSize, Style as CellStyle, Cell, Worksheet, Location } from "exceljs";
import { SheetExcelOption, SheetSection } from "./common-type.js";

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
  formula?: Cell["formula"];
};

export type SheetReportOptions = SheetExcelOption & {
  pageSize?: PaperSize;
  merges?: Record<string, { model: Location }>;
  cells: CellReportOptions[];
  columnWidths?: (number | undefined)[];
  rowHeights: Record<string, number>;
};

export type TableReportOptions = {
  sheets: SheetReportOptions[];
  name: string;
};
