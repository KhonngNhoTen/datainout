import * as path from "path";
import * as fs from "fs";

export type InoutConfigOptions = {
  import?: {
    outDir?: string;
    importDir?: string;
    dateFormat?: string;
  };
  report?: {
    templateDir?: string;
    reportDir?: string;
    excelSampleDir?: string;
  };
};

let globalConfig: InoutConfigOptions | undefined = undefined;

export function getConfig() {
  if (!globalConfig) {
    const configPath = path.join(process.cwd(), "inoutjs.config.json");
    if (fs.existsSync(configPath)) globalConfig = require(configPath);
    else globalConfig = {};
  }

  return globalConfig;
}
