import { SheetSection, TableData } from "./common-type.js";

export type FilterImportHandler = {
  sheetIndex: number;
  sheetName?: string;
  section: SheetSection;
};

export type ImporterBaseReaderType = "excel" | "csv";
export type ImporterBaseReaderStreamType = "excel-stream";
export type ImporterReaderType = ImporterBaseReaderType | ImporterBaseReaderStreamType;

export type ImporterHandlerFunction = (data: TableData, filter: FilterImportHandler) => Promise<any>;
