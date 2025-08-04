import { TemplateGenerator } from "../TemplateGenerator.js";
import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellReportOptions, SheetReportOptions } from "../../common/types/report-template.type.js";
import { getConfig } from "../../helpers/datainout-config.js";
import { pathReport } from "../../helpers/path-file.js";
import { ReaderExceljsHelper } from "../../helpers/excel.helper.js";
import { CellDataHelper, RowDataHelper, SheetDataHelper } from "../../common/types/excel-reader-helper.type.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { TableExcelOptions } from "../../common/types/common-type.js";

export class ExcelTemplateReport extends TemplateGenerator {
  private excelReaderHelper: ReaderExceljsHelper;
  private excelContent: TableExcelOptions<SheetReportOptions> = {
    sheets: [],
    name: "",
  };
  private currentSheet: SheetReportOptions = { cells: [], rowHeights: [] } as any;
  private useStyle?: boolean = true;

  constructor(template: string, useStyle: boolean = true) {
    super(template, pathReport);
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: async (data) => await this.onSheet(data),
      onCell: async (data) => await this.onCell(data),
      onRow: async (data) => await this.onRow(data),
      templateManager: new ExcelTemplateManager(),
      isSampleExcel: true,
    });
    this.useStyle = useStyle;
  }

  private async onCell(cell: CellDataHelper) {
    if ((cell.detail as any)._value.model.type !== exceljs.ValueType.Merge) {
      const fullAddress = cell.detail.fullAddress;
      if (cell.section === "footer") fullAddress.row = fullAddress.row - (cell.endTableAt ?? 0);
      const cellFormat: CellReportOptions = {
        address: cell.address,
        isVariable: cell.isVariable,
        value: { fieldName: cell.variableValue?.fieldName, hardValue: cell.label },
        section: cell.section,
        fullAddress,
        formula: cell.formula,
        keyName: cell?.variableValue?.fieldName ?? cell.label ?? "",
        index: this.currentSheet.cells.length + 1,
      } as any;
      if (this.useStyle) {
        cellFormat.style = cell.detail.style;
      }
      this.currentSheet.cells.push(cellFormat);
    }
  }

  private async onRow(row: RowDataHelper) {
    if (row.beginTableAt && row.rowIndex < row.beginTableAt) this.currentSheet.rowHeights[row.rowIndex] = row.detail.height;
    // if (row.endTableAt && row.rowIndex > row.endTableAt) this.currentSheet.rowHeights[row.rowIndex] = row.detail.height;
  }

  private async onSheet(sheet: SheetDataHelper) {
    const numberOfColumTable = this.currentSheet.cells.filter((e) => e.section === "table").length;
    if (this.useStyle)
      for (let i = 0; i < numberOfColumTable; i++) {
        if (!this.currentSheet.columnWidths) this.currentSheet.columnWidths = [];
        this.currentSheet.columnWidths.push(sheet.detail.getColumn(i + 1).width);
      }
    if (this.useStyle) this.currentSheet.merges = (sheet.detail as any)._merges;
    this.currentSheet.beginTableAt = sheet.beginTableAt;
    this.currentSheet.endTableAt = sheet.endTableAt;
    this.currentSheet.sheetIndex = sheet.sheetIndex;
    this.currentSheet.sheetName = sheet.name;
    this.currentSheet.keyTableAt = sheet.columnIndex;

    this.excelContent.sheets.push(this.currentSheet);
    this.currentSheet = {} as any;
  }
  generate(buffer: Buffer): Promise<any>;
  generate(fileSample: string): Promise<any>;
  async generate(arg: unknown) {
    if (arg instanceof Buffer) await this.excelReaderHelper.load(arg);
    else await this.excelReaderHelper.load(pathReport(arg + "", "layoutDir"));
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
