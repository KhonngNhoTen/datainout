import { PaperSize, Style as CellStyle, Cell, Location } from "exceljs";
import { BaseAttribute, SheetExcelOption, SheetSection } from "./common-type.js";
import { ExporterOutputType, ExporterStreamOutputType } from "./exporter.type.js";
import { ExcelTemplateManager } from "../core/Template.js";
/** Excel report template */
export type CellReportOptions = {
  address: string;
  fullAddress: Cell["fullAddress"];
  value: {
    hardValue?: any;
    fieldName?: string;
  };
  style: Partial<CellStyle>;
  isVariable: boolean;
  formula?: Cell["formula"];
  formatValue?: (data: any) => any;
} & BaseAttribute;

export type SheetReportOptions = SheetExcelOption<CellReportOptions> & {
  pageSize?: PaperSize;
  merges?: Record<string, { model: Location }>;
  columnWidths?: (number | undefined)[];
  rowHeights: Record<string, number>;
};

export type ReportOptions = {
  onError?: (data: any) => void;
  chunkSize?: number;
  useStyle?: boolean;
};

export type ReportStreamOptions = {
  useStyles?: boolean;
  useSharedStrings?: boolean;
  workerSize?: number;
  sleepTime?: number;
};
