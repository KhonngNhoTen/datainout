import * as path from "path";
import * as fs from "fs/promises";
import { ICommandAction } from "./ICommandAction.js";

export class InitializeAction implements ICommandAction {
  async handleAction(data: any) {
    if (data?.type === "js") this.genJsFile();
    else if (data?.type === "ts") this.genTsFile();
    else throw new Error(`File extension ${data?.type} is not supports.`);
  }

  genJsFile() {
    const _path = path.join(process.cwd(), "datainout.config.js");
    fs.writeFile(
      _path,
      `/** @type {import("datainout").DataInoutConfigOptions} */
module.exports = {
  dateFormat: "DD-MM-YYYY hh:mm:ss",
  templateExtension: ".js",
};`
    );
  }

  genTsFile() {
    const _path = path.join(process.cwd(), "datainout.config.ts");
    fs.writeFile(
      _path,
      `import { DataInoutConfigOptions } from 'datainout/types' 
export default {
  dateFormat: "DD-MM-YYYY hh:mm:ss",
  templateExtension: ".ts",
} as DataInoutConfigOptions;`
    );
  }
}
