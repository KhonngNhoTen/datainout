import { PaperSize, Style as CellStyle, Cell, Location } from "exceljs";
import { SheetExcelOption, SheetSection } from "./common-type.js";
import { ExporterOutputType, ExporterStreamOutputType } from "./exporter.type.js";
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
  formatValue?: (data: any) => any;
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

export type ReportOptions = {
  onError?: (data: any) => void;
  additionalCell?: CellReportOptions[];
  type?: ExporterOutputType;
};

export type ReportStreamOptions = {
  additionalCell?: CellReportOptions[];
  type?: ExporterStreamOutputType;
  useStyles?: boolean;
  useSharedStrings?: boolean;
  sleepTime?: number;
};
