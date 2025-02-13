import * as path from "path";
import { pathReport } from "../../helper/path-file.js";
import { getConfig } from "../../datainout-config.js";
export abstract class TemplateGenerator {
  protected templatePath: string;

  constructor(templatePath: string) {
    // template = `${template}.template.js`;

    // this.file = pathReport(file, "excelSampleDir");

    const paths = templatePath.replace("\\", "/").split("/");
    paths[paths.length - 1] = `${Date.now()}_${paths[paths.length - 1]}.retemplate.${
      getConfig()?.templateExtension === ".js" ? "js" : "ts"
    }`;
    templatePath = paths.length === 1 ? paths[0] : paths.reduce((init, val, i) => (i === 0 ? val : path.join(init, val)), "");

    this.templatePath = pathReport(templatePath ?? "", "templateDir");
  }

  abstract generate(buffer: Buffer): Promise<any>;
  abstract generate(fileSample: string): Promise<any>;
  abstract generate(arg: unknown): Promise<any>;
}
