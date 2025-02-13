import { PaperSize, Style as CellStyle, Cell, Worksheet, Location } from "exceljs";
import { SheetSection } from "../importers/type.js";
import { Exporter } from "./exporters/Exporter.js";
import { ReportDataIterator } from "./ReportDataIterator.js";
import { TableData } from "../type.js";

export type CellFormat = {
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

export type SheetFormat = {
  pageSize?: PaperSize;
  merges?: Record<string, { model: Location }>;
  beginTableAt: number;
  endTableAt?: number;
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

export type ExporterFactory = (type: ExporterList, templatePath: string, opts?: any) => Exporter;

export type CreateStreamOpts = {
  sheetBegin?: () => void;
  sheetFinish?: () => void;
  finish?: () => void;
  error?: (err: Error) => void;

  useStyles?: boolean;
  data: Array<Omit<TableData, "table"> & { table: ReportDataIterator; sheetName?: string }>;
};
