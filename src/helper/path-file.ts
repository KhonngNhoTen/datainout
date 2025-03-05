import * as path from "path";
import { getConfig } from "../datainout-config.js";
import { DataInoutConfigOptions, ListOfPathImports, ListOfPathReports } from "../type.js";

const config = getConfig();

export function pathReport(_path: string, fieldName?: ListOfPathReports) {
  if (!config?.report || !fieldName || !config?.report?.[fieldName]) return path.join(process.cwd(), _path);
  return path.join(process.cwd(), config.report[fieldName], _path);
}

export function pathImport(_path: string, fieldName?: ListOfPathImports) {
  if (!config?.import || !fieldName || !config?.import?.[fieldName]) return path.join(process.cwd(), _path);
  return path.join(process.cwd(), config.import[fieldName], _path);
}
