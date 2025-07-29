import { TypeParser } from "../../helpers/parse-type.js";
import { BaseReader } from "../../importers/readers/BaseReader.js";
import { SheetImportOptions } from "./import-template.type.js";
import { ImporterReaderType } from "./importer.type.js";

export type BaseReaderOptions = {
  type: ImporterReaderType;
  chunkSize?: number;
  typeParser?: TypeParser;
  jobCount?: number;
};

export type ReaderFactoryItem = {
  reader: BaseReader;
  isDefault?: boolean;
  name: string;
};
