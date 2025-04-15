import { pathImport, pathReport } from "../../helpers/path-file.js";
import { ExcelTemplateImport } from "../../template-generators/importer/excel.template.js";
import { ExcelTemplateReport } from "../../template-generators/reporter/excel.template.js";
import { ICommandAction } from "./ICommandAction.js";

export class GenerateAction implements ICommandAction {
  async handleAction(schema: string, options: any, ...args: any[]) {
    if (schema !== "import" && schema !== "report") throw new Error("Schema must be 'import' or 'report'!!");
    if (schema === "import") await this.genImportTemplate(options.nameTemplate, options.nameSource);
    else if (schema === "report") await this.genReportTemplate(options.nameTemplate, options.nameSource);
  }

  async genImportTemplate(templatePath: string, sourcePath: string) {
    await new ExcelTemplateImport(templatePath).generate(sourcePath);
  }

  async genReportTemplate(templatePath: string, sourcePath: string) {
    await new ExcelTemplateReport(templatePath).generate(sourcePath);
  }
}
