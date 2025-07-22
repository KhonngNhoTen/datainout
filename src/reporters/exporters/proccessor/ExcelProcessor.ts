import * as exceljs from "exceljs";
import { Writable } from "stream";
import { CellReportOptions } from "../../../common/types/report-template.type.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { EventRegister } from "../../../common/core/ListEvents.js";

type ExcelProcessorOptions = {
  workBook: exceljs.Workbook;
  template: ExcelTemplateManager<CellReportOptions>;
  event: EventRegister;
  header?: any;
  style?: "no-style" | "no-style-no-header" | "use-style";
  footer?: any;
};

export class ExcelProcessor {
  protected template: ExcelTemplateManager<CellReportOptions> = {} as any;
  protected workBook: exceljs.Workbook;
  protected headerData?: any;
  protected footerData?: any;
  protected event: EventRegister;
  protected style: "no-style" | "no-style-no-header" | "use-style" = "use-style";
  private titlesTable: string[] = [];
  private columnKeys?: { header: string; key: string }[] = [];

  constructor(opts: ExcelProcessorOptions) {
    this.workBook = opts.workBook;
    this.headerData = opts.header;
    this.footerData = opts.footer;
    this.template = opts.template;
    this.style = opts.style ?? "use-style";
    this.event = opts.event;
    this.titlesTable = this.template.GroupCells.table.map((e) => e.value.fieldName ?? "");
    this.columnKeys = opts.style !== "no-style-no-header" ? undefined : this.createColumnKey();
    console.log(this.style);
  }

  private createColumnKey() {
    const tableCells = this.template.GroupCells.table;
    const columns = this.template.GroupCells.header.map((e, i) => ({
      header: e.value.hardValue,
      key: tableCells[i].value.fieldName ?? "",
    }));
    return columns;
  }

  protected getOrCreateWorksheet(name: string) {
    return this.workBook.getWorksheet(name) ?? this.workBook.addWorksheet(name);
  }

  protected setHeader(headerData: any, sheet: exceljs.Worksheet): void;
  protected setHeader(headerData: any, sheetName: string): void;
  protected setHeader(headerData: any, arg: unknown) {
    if (!this.template.GroupCells.header) return;
    const workSheet = typeof arg === "string" ? this.getOrCreateWorksheet(arg) : (arg as exceljs.Worksheet);
    for (let i = 1; i <= this.template.SheetInformation.beginTableAt; i++) {
      const formats = this.template.GroupCells.header.filter((e) => e.fullAddress.row === i);
      //   formats.forEach((format) => this.setCell(headerData ?? {}, format, row.getCell(format.fullAddress.col)));
      this.addRow(headerData, workSheet, formats);
    }

    this.event.emitEvent("header", workSheet.name);
  }

  pushData(sheetName: string, batches: any[] | any, isFinish: boolean = false) {
    if (!this.workBook.getWorksheet(sheetName)) {
      this.event.emitEvent("begin", sheetName);
      const workSheet = this.workBook.addWorksheet(sheetName);
      if (this.columnKeys) workSheet.columns = this.columnKeys;
      else this.setHeader(this.headerData, workSheet);
    }
    const workSheet = this.getOrCreateWorksheet(sheetName);

    if (isFinish) {
      this.setFooter(this.footerData, workSheet);
      this.finalizeWorksheet(sheetName);
      return;
    }
    if (Array.isArray(batches)) {
      for (let i = 0; i < batches.length; i++) {
        if (this.style === "no-style-no-header") workSheet.addRow(batches[i]).commit();
        else if (this.style === "use-style") this.addRow(batches[i], workSheet, this.template.GroupCells.table);
        else this.addRowWithoutStyle(batches[i], workSheet);
      }
    } else if (batches !== null) this.addRow(batches, workSheet, this.template.GroupCells.table);
  }

  protected setFooter(footerData: any, sheetName: string): void;
  protected setFooter(footerData: any, sheet: exceljs.Worksheet): void;
  protected setFooter(footerData: any, arg: unknown): void {
    const workSheet = typeof arg === "string" ? this.getOrCreateWorksheet(arg) : (arg as exceljs.Worksheet);

    if (this.template.GroupCells.footer) {
      const groupFooter = this.template.GroupCells.footer
        .sort((a, b) => b.fullAddress.row - a.fullAddress.row)
        .reduce((acc, value) => {
          const row = value.fullAddress.row;
          acc[row] = acc[row] ? [...acc[row], value] : [value];
          return acc;
        }, {} as Record<number, CellReportOptions[]>);

      for (const [rowIndex, footers] of Object.entries(groupFooter)) {
        this.addRow(footerData, workSheet, footers);
      }
    }

    if (this.style === "use-style") this.mergeCells(workSheet);
    if (this.style === "use-style") this.setWidthsAndHeights(workSheet);
    this.event.emitEvent("footer", workSheet.name);
  }

  finalizeWorksheet(sheetName: string) {
    this.event.emitEvent("end", sheetName);
  }

  protected addRowWithoutStyle(rowdata: any, workSheet: exceljs.Worksheet) {
    // const row = workSheet.addRow([]);
    // for (let i = 0; i < this.titlesTable.length; i++) {
    //   row.getCell(i + 1).value = rowdata[this.titlesTable[i]];
    // }
    // row.commit();
    const row = [];
    for (let i = 0; i < this.titlesTable.length; i++) {
      row.push(rowdata[this.titlesTable[i]]);
    }
    workSheet.addRow(row).commit();
  }

  protected addRow(data: any, workSheet: exceljs.Worksheet, cellDesc: CellReportOptions[]): exceljs.Row {
    const row = workSheet.addRow([]);
    for (let i = 0; i < cellDesc.length; i++) {
      const cellsOpt = cellDesc[i];
      const cell = row.getCell(cellsOpt.fullAddress.col);
      this.setCell(data, cellsOpt, cell);
    }
    return row;
  }

  protected setCell(rowData: any, cellOpt: CellReportOptions, cell: exceljs.Cell): exceljs.Cell {
    // console.time("setCell");
    if (this.style === "use-style") cell.style = cellOpt.style;
    if (cellOpt.formula) {
      cell.value = {
        formula: cellOpt.formula,
      };
    } else {
      let value = cellOpt.isVariable ? rowData[(cellOpt.value as any).fieldName] : (cellOpt.value as any).hardValue;
      if (cellOpt.formatValue) value = cellOpt.formatValue(value);
      cell.value = value;
    }
    // console.timeEnd("setCell");

    return cell;
  }

  protected mergeCells(sheet: exceljs.Worksheet) {
    const merges = (this.template.SheetTemplate as any).merges;
    if (merges) {
      Object.keys(merges).forEach((masterCell: string) => {
        const { top, left, right, bottom } = merges[masterCell].model;
        sheet.mergeCells(top, left, bottom, right);
      });
    }
  }

  protected setWidthsAndHeights(sheet: exceljs.Worksheet) {
    const columnWidths = (this.template.SheetTemplate as any).columnWidths;
    const rowHeights = (this.template.SheetTemplate as any).rowHeights;

    // Set column's width
    columnWidths?.forEach((colW: any, i: number) => {
      if (sheet?.columns && sheet?.columns[i] && colW) sheet.columns[i].width = colW;
    });

    // Set header height
    if (rowHeights)
      Object.keys(rowHeights).forEach((rowIndex, i) => {
        if (rowHeights[rowIndex]) sheet.getRow(rowHeights[i]).height = rowHeights[rowIndex];
      });
  }
}

export class ExcelStreamProcessor extends ExcelProcessor {
  override finalizeWorksheet(sheetName: string): void {
    const worksheet = this.getOrCreateWorksheet(sheetName);
    worksheet.commit();
    this.event.emitEvent("end", sheetName);
  }

  protected override addRow(data: any, workSheet: exceljs.Worksheet, cellDesc: CellReportOptions[]): exceljs.Row {
    // console.time("addRow");
    const row = super.addRow(data, workSheet, cellDesc);
    // console.timeEnd("addRow");
    // console.time("commit-row");

    if (this.style !== "use-style") {
      row.commit();
    }
    // console.timeEnd("commit-row");

    return row;
  }

  async finalizeWorkbook() {
    await (this.workBook as exceljs.stream.xlsx.WorkbookWriter).commit();
    this.event.emitEvent("finish");
  }
}
