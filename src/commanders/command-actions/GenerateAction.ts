import { TemplateGenerator } from "../../template-mappers/template-generator.js";
import { ICommandAction } from "./ICommandAction.js";

export class GenerateAction implements ICommandAction {
  async handleAction(schema: string, options: any, ...args: any[]) {
    if (schema !== "import" && schema !== "report") throw new Error("Schema must be 'import' or 'report'!!");
    if (schema === "import") await this.genImportTemplate(options.nameTemplate, options.nameSource);
    else if (schema === "report") await this.genReportTemplate(options.nameTemplate, options.nameSource);
  }

  async genImportTemplate(templatePath: string, sourcePath: string) {
    await TemplateGenerator.create(templatePath,sourcePath, "import")
  }

  async genReportTemplate(templatePath: string, sourcePath: string) {
    await TemplateGenerator.create(templatePath,sourcePath, "report")
  }
}
