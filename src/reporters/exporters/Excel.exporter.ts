import * as exceljs from "exceljs";
import { PassThrough } from "stream";
import { Exporter } from "./Exporter.js";
import { CellFormat, CreateStreamOpts, ExcelFormat, ReportData, SheetFormat } from "../type.js";
import { getFileExtension } from "../../helper/get-file-extension.js";
import { WriterStreanm } from "./stream/WriterStream.js";
import { ExcelWriterStream } from "./stream/ExcelWriterStream.js";
export class ExcelExporter extends Exporter {
  excelFormat: ExcelFormat = [];
  workBook?: exceljs.Workbook;

  constructor(templatePath: string, opts: any) {
    super(templatePath, opts);
    this.excelFormat =
      getFileExtension(this.templatePath) === "js"
        ? (require(templatePath) as ExcelFormat)
        : (require(templatePath).default as ExcelFormat);
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
    const titleTables = excelFormat.cellFomats.filter((e) => e.section === "table" && !e.isVariable);
    titleTables.forEach((titleTable) => {
      this.createCell(sheet, titleTable, undefined);
    });

    const contentTables = excelFormat.cellFomats.filter((e) => e.section === "table" && e.isVariable);
    reportData.table?.forEach((rowData, i) => {
      const row = sheet.addRow([]);
      contentTables.forEach((e) => {
        const cell = row.getCell(e.fullAddress.col);
        cell.style = e.style;
        cell.value = rowData[e.value?.fieldName ?? ""];
      });
    });

    // Add cells in footer-section
    sheet = this.mergesCells(sheet, excelFormat);
    sheet = this.setWidthsAndHeights(sheet, excelFormat);
  }

  protected createCell(sheet: exceljs.Worksheet, cellFormat: CellFormat, cellValue: any, address?: string) {
    const cell = sheet.getCell(address ?? cellFormat.address);
    cell.style = cellFormat.style;
    cell.value = !cellFormat.isVariable ? cellFormat.value.hardValue : cellValue;
  }

  protected mergesCells(sheet: exceljs.Worksheet, sheetFormat: SheetFormat) {
    if (sheetFormat.merges) {
      const merges = sheetFormat.merges;
      Object.keys(sheetFormat.merges).forEach((masterCell: string) => {
        const { top, left, right, bottom } = merges[masterCell].model;
        sheet.mergeCells(top, left, bottom, right);
      });
    }

    return sheet;
  }

  protected setWidthsAndHeights(sheet: exceljs.Worksheet, sheetFormat: SheetFormat) {
    // Set column's width
    sheetFormat.columnWidths?.forEach((colW, i) => {
      if (sheet.columns[i]) sheet.columns[i].width = colW;
    });

    // Set header and footer height
    const rowHeights = sheetFormat.rowHeights;
    Object.keys(rowHeights).forEach((rowIndex) => {
      if (rowHeights[rowIndex]) sheet.getRow(rowHeights[rowIndex]).height = rowHeights[rowIndex];
    });

    return sheet;
  }

  writerStream(opts: CreateStreamOpts): WriterStreanm {
    return new ExcelWriterStream(opts, this.excelFormat);
  }
}
