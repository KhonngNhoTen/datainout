import { TypeParser } from "../../helpers/parse-type.js";
import { EventType } from "./common-type.js";

export type ConvertorRows2TableDataOpts = {
  chunkSize?: number;
  typeParser?: TypeParser;
  onError?: EventType["error"];
};

export type GroupValueRow = Record<string, any>;
