import * as path from "path";
import { pathImport, pathReport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
export abstract class TemplateGenerator {
  protected templatePath: string;
  protected functionPathFormat: typeof pathReport | typeof pathImport;

  constructor(templatePath: string, functionPathFormat: typeof pathReport | typeof pathImport) {
    this.functionPathFormat = functionPathFormat;
    this.templatePath = this.functionPathFormat(templatePath ?? "", "templateDir");
    this.templatePath = `${this.templatePath}${getConfig()?.templateExtension === ".ts" ? ".ts" : ".js"}`;
  }

  abstract generate(buffer: Buffer): Promise<any>;
  abstract generate(fileSample: string): Promise<any>;
  abstract generate(arg: unknown): Promise<any>;
}
