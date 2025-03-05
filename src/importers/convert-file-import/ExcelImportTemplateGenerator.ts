import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellDescription, CellType, ImportFileDesciptionOptions, SheetDesciptionOptions, SheetSection } from "../type";
import { pathImport } from "../../helper/path-file";
import { CellDataHelper, ExcelReaderHelper, SheetDataHelper } from "../../helper/excel-reader-helper";
import path from "path";
import { getConfig } from "../../datainout-config";
import { ImportFileDesciption } from "../reader/ImporterFileDescription";

export class ExcelImportTemplateGenerator {
  private templatePath: string;
  private excelReaderHelper: ExcelReaderHelper;
  private importDesc: ImportFileDesciptionOptions = { sheets: [] };
  private currentSheet: CellDescription[] = [];

  constructor(templatePath: string) {
    const paths = templatePath.replace("\\", "/").split("/");
    paths[paths.length - 1] = `${Date.now()}_${paths[paths.length - 1]}.imtemplate.${
      getConfig()?.templateExtension === ".js" ? "js" : "ts"
    }`;
    templatePath = paths.length === 1 ? paths[0] : paths.reduce((init, val, i) => (i === 0 ? val : path.join(init, val)), "");

    this.templatePath = pathImport(templatePath ?? "", "templateDir");

    this.excelReaderHelper = new ExcelReaderHelper({
      onCell: (data) => this.onCell(data),
      onSheet: (data) => this.onSheet(data),
    });
  }

  async write(file: string): Promise<any>;
  async write(buffer: Buffer): Promise<any>;
  async write(arg: unknown): Promise<any> {
    console.log(`Create template import file: [${this.templatePath}]`);

    if (!(arg instanceof Buffer)) {
      arg = pathImport(arg as string, "sampleFileDir");
      arg = fs.readFile(arg as string);
    }
    await this.excelReaderHelper.load(arg as Buffer);

    await fs.writeFile(this.templatePath, this.genContentFile(this.importDesc), "utf-8");
    console.log(`Create file successfully!`);
  }

  private onCell(cell: CellDataHelper) {
    if (!cell.isVariable || (cell.detail as any)._value.model.type === exceljs.ValueType.Merge) return;
    this.currentSheet.push({
      fieldName: cell.variableValue?.fieldName ?? "",
      section: cell.section,
      type: cell.variableValue?.type ?? "string",
      address: cell.address,
    });
  }

  private onSheet(sheet: SheetDataHelper) {
    if (this.currentSheet.length > 0) {
      this.importDesc.sheets.push({
        content: this.currentSheet,
        endTableAt: sheet.endTableAtAt,
        name: sheet.name,
        beginTableAt: sheet.beginTableAt,
      });
      this.currentSheet = [];
    }
  }

  private genContentFile(importDesciption: ImportFileDesciptionOptions): string {
    if (getConfig()?.templateExtension === ".js")
      return `/** @type {import("datainout").ImportFileDesciptionOptions} */
const template = ${JSON.stringify(importDesciption, null, undefined)};
module.exports = template;`;
    return `import { ImportFileDesciptionOptions } from "datainout";
const template : ImportFileDesciptionOptions = ${JSON.stringify(importDesciption, null, undefined)};    
export default template`;
  }
}
