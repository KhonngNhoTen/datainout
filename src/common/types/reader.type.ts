import { TypeParser } from "../../helpers/parse-type.js";
import { BaseReader } from "../../importers/readers/BaseReader.js";
import { ExcelTemplateManager } from "../core/Template.js";
import { CellImportOptions, SheetImportOptions } from "./import-template.type.js";
import { ImporterReaderType } from "./importer.type.js";

export type BaseReaderOptions = {
  type: ImporterReaderType;
  chunkSize?: number;
  typeParser?: TypeParser;
  jobCount?: number;
  templateManager: ExcelTemplateManager<CellImportOptions>;
};

export type ReaderFactoryItem = {
  reader: BaseReader;
  isDefault?: boolean;
  name: string;
};
