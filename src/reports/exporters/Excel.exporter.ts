import * as exceljs from "exceljs";
import { Exporter } from "./Exporter";
import { CellFormat, ExcelFormat, ReportData } from "../type";
export class ExcelExporter implements Exporter {
  excelFormat: ExcelFormat = [];
  workBook?: exceljs.Workbook;

  setup(templatePath: string) {
    this.excelFormat = require(templatePath) as ExcelFormat;
    this.workBook = new exceljs.Workbook();
  }

  async writeFile(reportDatas: ReportData | ReportData[], path: string) {
    if (!this.workBook) throw new Error("WorkBook is Null");
    await this.createContent(reportDatas);
    this.workBook.xlsx.writeFile(path);
  }

  async buffer(reportDatas: ReportData | ReportData[]): Promise<Buffer> {
    if (!this.workBook) throw new Error("WorkBook is Null");
    await this.createContent(reportDatas);
    const buffer = await this.workBook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  private async createContent(reportDatas: ReportData | ReportData[]) {
    if (!Array.isArray(reportDatas)) reportDatas = [reportDatas];
    for (let i = 0; i < reportDatas.length; i++) {
      const reportData = reportDatas[i];
      const workSheet = this.workBook?.addWorksheet();
      if (!workSheet) continue;
      this.createSheet(workSheet, reportData, i);
    }
  }

  private createSheet(sheet: exceljs.Worksheet, reportData: ReportData, sheetIndex: number) {
    const excelFormat = this.excelFormat[sheetIndex];
    // Add cells in header-section
    const headerCells = excelFormat.cellFomats.filter((e) => e.section === "header");
    headerCells.forEach((headerCell) => {
      this.createCell(sheet, headerCell, (reportData.header as any)[headerCell?.value?.fieldName ?? ""]);
    });

    // Add cells in table-section
    const titleTables = excelFormat.cellFomats.filter((e) => e.section === "table" && e.isHardCell);
    titleTables.forEach((titleTable) => {
      this.createCell(sheet, titleTable, undefined);
    });

    const contentTables = excelFormat.cellFomats.filter((e) => e.section === "table" && !e.isHardCell);
    let isFirstRow = true;
    reportData.table?.forEach((row, i) => {
      // is first row in table
      if (isFirstRow) {
        contentTables.forEach((contentTable) => {
          this.createCell(
            sheet,
            contentTable,
            (reportData.table as any)[i][contentTable?.value?.fieldName ?? ""],
            `${contentTable.address}${excelFormat.beginTable + 1}`,
          );
        });
        isFirstRow = false;
      } else {
        const rowValues = Object.keys(row).map((cell) => row[cell]);
        sheet.addRow(rowValues, "i");
      }
    });

    // Add cells in footer-section

    // Set column's width
    excelFormat.columnWidths?.forEach((colW, i) => {
      sheet.columns[i].width = colW;
    });

    // Set header and footer height
    const rowHeights = excelFormat.rowHeights;
    Object.keys(rowHeights).forEach((rowIndex) => (sheet.getRow(rowHeights[rowIndex]).height = rowHeights[rowIndex]));
  }

  private createCell(sheet: exceljs.Worksheet, cellFormat: CellFormat, cellValue: any, address?: string) {
    const cell = sheet.getCell(address ?? cellFormat.address);
    cell.style = cellFormat.style;
    cell.value = cellFormat.isHardCell ? cellFormat.value.hardValue : cellValue;
  }
}
