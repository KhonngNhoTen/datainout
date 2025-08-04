import { PartialDataTransfer } from "../../reporters/PartialDataTransfer.js";
import { ExcelTemplateManager } from "../core/Template.js";
import { CellReportOptions, ReportOptions, ReportStreamOptions, SheetReportOptions } from "./report-template.type.js";

export type ExporterOutputType = "csv" | "excel" | "html" | "pdf";
export type ExporterStreamOutputType = "excel";
export type ExporterMethodType = "full-load" | "stream";

export type ExporterOptions = {
  templateManager?: ExcelTemplateManager<CellReportOptions>;
  templatePath?: string;
  action: "write" | "buffer" | "stream";
  reportPath?: string;
} & Partial<ReportOptions>;

export type ExporterStreamOptions = {
  content: { header?: any; footer?: any; table: PartialDataTransfer };
  templateManager?: ExcelTemplateManager<CellReportOptions>;
} & Omit<ExporterOptions, "templateManager"> &
  ReportStreamOptions;
