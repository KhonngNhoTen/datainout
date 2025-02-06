import { PaperSize, Style as CellStyle } from "exceljs";
import { ResultOfImport, SheetSection } from "../type";

export type CellFormat = {
  address: string;
  section?: SheetSection;
  value: {
    hardValue?: any;
    fieldName?: string;
  };
  style: Partial<CellStyle>;
  isHardCell: boolean;
};

export type SheetFormat = {
  pageSize?: PaperSize;
  beginTable: number;
  endTable?: number;
  cellFomats: CellFormat[];
  columnWidths?: (number | undefined)[];
  rowHeights: Record<string, number>;
};

export type ReportData = {
  header?: {};
  footer?: {};
  table?: any[];
};

export type ExporterList = "excel" | "pdf" | "html";

export type ExcelFormat = SheetFormat[];
