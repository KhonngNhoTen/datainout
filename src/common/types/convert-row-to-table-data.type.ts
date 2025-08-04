import { TypeParser } from "../../helpers/parse-type.js";
import { ExcelTemplateManager } from "../core/Template.js";
import { EventType } from "./common-type.js";
import { CellImportOptions } from "./import-template.type.js";

export type ConvertorRows2TableDataOpts = {
  chunkSize?: number;
  typeParser?: TypeParser;
  onError?: EventType["error"];
  templateManager?: ExcelTemplateManager<CellImportOptions>;
};

export type GroupValueRow = Record<string, any>;
