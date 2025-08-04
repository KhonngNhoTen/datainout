import { ImporterHandler } from "../../importers/ImporterHandler.js";
import { SheetSection, TableData } from "./common-type.js";

export type FilterImportHandler = {
  sheetIndex: number;
  sheetName?: string;
  section: SheetSection;
  isHasNext: boolean;
};

export type ImporterBaseReaderType = "excel" | "csv";
export type ImporterBaseReaderStreamType = "excel-stream";
export type ImporterReaderType = ImporterBaseReaderType | ImporterBaseReaderStreamType;

export type ImporterHandlerFunctionData = TableData | Error | Error[];
export type ImporterHandlerFunction = (
  data: ImporterHandlerFunctionData,
  filter: FilterImportHandler
) => Promise<ImporterHandlerFunctionData>;
export type ImporterHandlerInstance = ImporterHandler<any> | ImporterHandlerFunction[];

export type ImporterLoadFunctionOpions = {
  type?: ImporterBaseReaderType;
  chunkSize?: number;
  ignoreErrors?: boolean;
  jobCount?: number;
};
