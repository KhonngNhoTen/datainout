const fs = require("fs");
const path = require("path");

export type InoutConfigOptions = {
  import: {
    outDir?: string;
    excelDir?: string;
    dateFormat?: string;
  };
};

class InoutConfig {
  options: InoutConfigOptions;
  constructor(opts: InoutConfigOptions) {
    this.options = opts;
  }
}

const configPath = path.join(process.cwd(), "inoutrc.json");
export const inoutConfig = new InoutConfig(require(configPath));
