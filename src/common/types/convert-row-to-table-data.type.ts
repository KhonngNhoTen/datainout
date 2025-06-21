import { TypeParser } from "../../helpers/parse-type.js";

export type ConvertorRows2TableDataOpts = {
  chunkSize?: number;
  typeParser?: TypeParser;
};

export type GroupValueRow = Record<string, any>;
