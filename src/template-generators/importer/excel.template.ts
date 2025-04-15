import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellImportOptions, TableImportOptions } from "../../common/types/import-template.type.js";
import { getConfig } from "../../helpers/datainout-config.js";
import { CellDataHelper, ExcelReaderHelper, SheetDataHelper } from "../../helpers/excel-reader-helper.js";
import { pathImport } from "../../helpers/path-file.js";
import { TemplateGenerator } from "../TemplateGenerator.js";

export class ExcelTemplateImport extends TemplateGenerator {
  private excelReaderHelper: ExcelReaderHelper;
  private excelContent: TableImportOptions = {
    sheets: [],
    name: "",
  };
  private currentSheet: CellImportOptions[] = [];

  constructor(templatePath: string) {
    super(templatePath, pathImport);

    this.excelReaderHelper = new ExcelReaderHelper({
      onCell: async (data: any) => this.onCell(data),
      onSheet: async (data: any) => this.onSheet(data),
    });
  }

  private onCell(cell: CellDataHelper) {
    if (!cell.isVariable || (cell.detail as any)._value.model.type === exceljs.ValueType.Merge) return;
    const fullAddress = cell.detail.fullAddress;
    if (cell.section === "footer") fullAddress.row = fullAddress.row - (cell.endTableAtAt ?? 0);
    this.currentSheet.push({
      keyName: cell.variableValue?.fieldName ?? "",
      section: cell.section,
      type: cell.variableValue?.type ?? "string",
      address: cell.address,
      fullAddress,
    });
  }

  private onSheet(sheet: SheetDataHelper) {
    if (this.currentSheet.length > 0) {
      this.excelContent.sheets.push({
        cells: this.currentSheet,
        endTableAt: sheet.endTableAtAt,
        sheetName: sheet.name,
        beginTableAt: sheet.beginTableAt,
        keyTableAt: sheet.columnIndex,
        sheetIndex: sheet.sheetIndex,
      });
      this.currentSheet = [];
    }
  }

  private genContentFile(excelContent: TableImportOptions): string {
    if (getConfig()?.templateExtension === ".js")
      return `/** @type {import("datainout").ImportFileDesciptionOptions} */
const template = ${JSON.stringify(excelContent, null, undefined)};
module.exports = template;`;
    return `import { ImportFileDesciptionOptions } from "datainout";
const template : ImportFileDesciptionOptions = ${JSON.stringify(excelContent, null, undefined)};    
export default template`;
  }

  generate(buffer: Buffer): Promise<any>;
  generate(fileSample: string): Promise<any>;
  async generate(arg: unknown) {
    console.log(`Create template import file: [${this.templatePath}]`);

    if (!(arg instanceof Buffer)) {
      arg = pathImport(arg as string, "excelSampleDir");
      arg = await fs.readFile(arg as string);
    }

    await this.excelReaderHelper.load(arg as Buffer);
    await fs.writeFile(this.templatePath, this.genContentFile(this.excelContent), "utf-8");
    console.log(`Create file successfully!`);
  }
}
