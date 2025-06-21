import { TypeParser } from "../../helpers/parse-type.js";
import { BaseReader } from "../../importers/readers/BaseReader.js";
import { FilterImportHandler, ImporterReaderType } from "./importer.type.js";

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
