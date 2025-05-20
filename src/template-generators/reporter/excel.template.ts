import { TemplateGenerator } from "../TemplateGenerator.js";
import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellDataHelper, ExcelReaderHelper, RowDataHelper, SheetDataHelper } from "../../helpers/excel-reader-helper.js";
import { CellReportOptions, TableReportOptions, SheetReportOptions } from "../../common/types/report-template.type.js";
import { getConfig } from "../../helpers/datainout-config.js";
import { pathReport } from "../../helpers/path-file.js";

export class ExcelTemplateReport extends TemplateGenerator {
  private excelReaderHelper: ExcelReaderHelper;
  private excelContent: TableReportOptions = {
    sheets: [],
    name: "",
  };
  private currentSheet: SheetReportOptions = {
    beginTableAt: 0,
    cells: [],
    rowHeights: {},
  };

  constructor(template: string) {
    super(template, pathReport);
    this.excelReaderHelper = new ExcelReaderHelper({
      onCell: async (cell) => await this.onCell(cell),
      onRow: async (data: any) => await this.onRow(data),
      onSheet: async (data: any) => await this.onSheet(data),
    });
  }

  private async onCell(cell: CellDataHelper) {
    if ((cell.detail as any)._value.model.type !== exceljs.ValueType.Merge) {
      const fullAddress = cell.detail.fullAddress;
      if (cell.section === "footer") fullAddress.row = fullAddress.row - (cell.endTableAtAt ?? 0);
      const cellFormat: CellReportOptions = {
        address: cell.address,
        isVariable: cell.isVariable,
        style: cell.detail.style,
        value: { fieldName: cell.variableValue?.fieldName, hardValue: cell.label },
        section: cell.section,
        fullAddress,
      };
      this.currentSheet.cells.push(cellFormat);
    }
  }

  private async onRow(row: RowDataHelper) {
    if (row.beginTableAt && row.rowIndex < row.beginTableAt) this.currentSheet.rowHeights[row.rowIndex] = row.detail.height;
    if (row.endTableAtAt && row.rowIndex > row.endTableAtAt) this.currentSheet.rowHeights[row.rowIndex] = row.detail.height;
  }

  private async onSheet(sheet: SheetDataHelper) {
    this.currentSheet.columnWidths = sheet.detail.columns.map((col) => col.width);
    this.currentSheet.beginTableAt = sheet.beginTableAt;
    this.currentSheet.endTableAt = sheet.endTableAtAt;
    this.currentSheet.merges = (sheet.detail as any)._merges;

    this.excelContent.sheets.push(this.currentSheet);
    this.currentSheet = { cells: [], beginTableAt: 1, rowHeights: {}, columnWidths: [] };
  }
  generate(buffer: Buffer): Promise<any>;
  generate(fileSample: string): Promise<any>;
  async generate(arg: unknown) {
    if (arg instanceof Buffer) await this.excelReaderHelper.load(arg);
    else await this.excelReaderHelper.load(pathReport(arg + "", "excelSampleDir"));
    let contentFile = "";
    if (getConfig()?.templateExtension === ".js")
      contentFile = `/** @type {import("inoutjs").ExcelFormat} */
const template = ${JSON.stringify(this.excelContent, null, undefined)};
module.exports = template;`;
    else
      contentFile = `import { ExcelFormat } from "datainout";
const template : ExcelFormat = ${JSON.stringify(this.excelContent, null, undefined)};    
export default template`;
    await fs.writeFile(this.templatePath, contentFile);
  }
}
