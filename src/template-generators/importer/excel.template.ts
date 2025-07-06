import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellImportOptions, SheetImportOptions } from "../../common/types/import-template.type.js";
import { getConfig } from "../../helpers/datainout-config.js";
import { pathImport } from "../../helpers/path-file.js";
import { TemplateGenerator } from "../TemplateGenerator.js";
import { ReaderExceljsHelper } from "../../helpers/excel.helper.js";
import { CellDataHelper, SheetDataHelper } from "../../common/types/excel-reader-helper.type.js";
import { SheetSection, TableExcelOptions } from "../../common/types/common-type.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";

export class ExcelTemplateImport extends TemplateGenerator {
  private excelReaderHelper: ReaderExceljsHelper;
  private excelContent: TableExcelOptions<SheetImportOptions> = {
    sheets: [],
    name: "",
  };
  private currentSheet: CellImportOptions[] = [];

  constructor(templatePath: string) {
    super(templatePath, pathImport);
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: async (data) => await this.onSheet(data),
      onCell: async (data) => await this.onCell(data),
      templateManager: new ExcelTemplateManager(),
      isSampleExcel: true,
    });
  }

  private async onCell(cell: CellDataHelper) {
    if (!cell.isVariable || (cell.detail as any)._value.model.type === exceljs.ValueType.Merge) return;
    const fullAddress = cell.detail.fullAddress;
    if (cell.section === "footer") fullAddress.row = fullAddress.row - (cell.endTableAt ?? 0);

    this.currentSheet.push({
      keyName: cell.variableValue?.fieldName ?? "",
      section: cell.section,
      type: cell.variableValue?.type ?? "string",
      address: cell.address,
      fullAddress,
    });
  }

  private async onSheet(sheet: SheetDataHelper) {
    if (this.currentSheet.length > 0) {
      this.excelContent.sheets.push({
        cells: this.currentSheet,
        endTableAt: sheet.endTableAt,
        sheetName: sheet.name,
        beginTableAt: sheet.beginTableAt,
        keyTableAt: sheet.columnIndex,
        sheetIndex: sheet.sheetIndex,
      });
      this.currentSheet = [];
    }
  }

  private genContentFile(excelContent: TableExcelOptions<SheetImportOptions>): string {
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
