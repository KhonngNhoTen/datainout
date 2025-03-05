import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellFormat, ExcelFormat, SheetFormat } from "../type.js";
import { TemplateGenerator } from "./TemplateGenerator.js";
import { CellDataHelper, ExcelReaderHelper, RowDataHelper, SheetDataHelper } from "../../helper/excel-reader-helper.js";
import { pathReport } from "../../helper/path-file.js";
import { getConfig } from "../../datainout-config.js";

/**
 * Convert excel file into ExcelFormat type and save on template path
 */
export class Excel2ExcelTemplateGenerator extends TemplateGenerator {
  private excelReaderHelper: ExcelReaderHelper;
  private excelFormat: ExcelFormat = [];
  private currentSheetFormat: SheetFormat = { cellFomats: [], beginTableAt: 1, rowHeights: {}, columnWidths: [] };

  constructor(template: string) {
    super(template);
    this.excelReaderHelper = new ExcelReaderHelper({
      onCell: (data) => this.onCell(data),
      onRow: (data) => this.onRow(data),
      onSheet: (data) => this.onSheet(data),
    });
  }

  private onCell(cell: CellDataHelper) {
    if ((cell.detail as any)._value.model.type !== exceljs.ValueType.Merge) {
      const cellFormat: CellFormat = {
        address: cell.address,
        isVariable: cell.isVariable,
        style: cell.detail.style,
        value: { fieldName: cell.variableValue?.fieldName, hardValue: cell.label },
        section: cell.section,
        fullAddress: cell.detail.fullAddress,
      };
      this.currentSheetFormat.cellFomats.push(cellFormat);
    }
  }

  private onRow(row: RowDataHelper) {
    if (row.beginTableAt && row.rowIndex < row.beginTableAt) this.currentSheetFormat.rowHeights[row.rowIndex] = row.detail.height;
    if (row.endTableAtAt && row.rowIndex > row.endTableAtAt) this.currentSheetFormat.rowHeights[row.rowIndex] = row.detail.height;
  }

  private onSheet(sheet: SheetDataHelper) {
    this.currentSheetFormat.columnWidths = sheet.detail.columns.map((col) => col.width);
    this.currentSheetFormat.beginTableAt = sheet.beginTableAt;
    this.currentSheetFormat.endTableAt = sheet.endTableAtAt;
    this.currentSheetFormat.merges = (sheet.detail as any)._merges;
    this.excelFormat.push(this.currentSheetFormat);
    this.currentSheetFormat = { cellFomats: [], beginTableAt: 1, rowHeights: {}, columnWidths: [] };
  }

  async generate(arg: string): Promise<void>;
  async generate(arg: Buffer): Promise<void>;
  async generate(arg: unknown): Promise<void> {
    if (arg instanceof Buffer) await this.excelReaderHelper.load(arg);
    else await this.excelReaderHelper.load(pathReport(arg + "", "excelSampleDir"));
    let contentFile = "";
    if (getConfig()?.templateExtension === ".js")
      contentFile = `/** @type {import("inoutjs").ExcelFormat} */
  const template = ${JSON.stringify(this.excelFormat, null, undefined)};
  module.exports = template;`;
    else
      contentFile = `import { ExcelFormat } from "datainout";
const template : ExcelFormat = ${JSON.stringify(this.excelFormat, null, undefined)};    
export default template`;
    await fs.writeFile(this.templatePath, contentFile);
  }
}
