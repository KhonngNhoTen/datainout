import { TypeParser } from "../../helpers/parse-type.js";
import { BaseReader } from "../../importers/readers/BaseReader.js";
import { FilterImportHandler, ImporterReaderType } from "./importer.type.js";

export type EventType = {
  rFinish: () => void;
  rBegin: (sheetName?: string) => void;
  rData: (filter?: FilterImportHandler) => void;
  rEnd: (sheetName?: string) => void;

  /** Handle error. Return false to cancel import, otherhands return true */
  rError: (error: Error) => boolean;
};

export type BaseReaderOptions = {
  type: ImporterReaderType;
  chunkSize?: number;
  typeParser?: TypeParser;
};

export type ReaderFactoryItem = {
  reader: BaseReader;
  isDefault?: boolean;
  name: string;
};
