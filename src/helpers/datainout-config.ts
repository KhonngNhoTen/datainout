import * as path from "path";
import * as fs from "fs";
import { DataInoutConfigOptions } from "../common/types/config.type.js";

let globalConfig: DataInoutConfigOptions = {};

export function getConfig() {
  if (Object.keys(globalConfig).length === 0) {
    const jsPath = path.join(process.cwd(), "datainout.config.js");
    const tsPath = path.join(process.cwd(), "datainout.config.ts");
    if (fs.existsSync(jsPath)) globalConfig = require(jsPath);
    else if (fs.existsSync(tsPath)) globalConfig = require(tsPath).default;
    else globalConfig = { templateExtension: ".js", dateFormat: "DD-MM-YYYY" };
  }

  return globalConfig;
}
