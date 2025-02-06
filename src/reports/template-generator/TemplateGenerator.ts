import { pathReport } from "../../helper/path-file";
export abstract class TemplateGenerator {
  protected file: string;
  protected template: string;

  constructor(file: string, template: string) {
    template = `${template}.template.js`;

    this.file = pathReport(file, "excelSampleDir");

    this.template = pathReport(template, "templateDir");
  }

  abstract generate(): Promise<any>;
}
