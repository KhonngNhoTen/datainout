import { PassThrough } from "stream";
import { WriterStreanm } from "./WriterStream.js";
import * as exceljs from "exceljs";
import { CellFormat, CreateStreamOpts, ExcelFormat, SheetFormat } from "../../type.js";

export class ExcelWriterStream implements WriterStreanm {
  private _workBookWriter: exceljs.stream.xlsx.WorkbookWriter;
  private _workSheet: exceljs.Worksheet[] = [];
  private _stream = new PassThrough();
  private _indexRow = 1;

  private _sheetBegin?: () => void;
  private _sheetFinish?: () => void;
  private _finish?: () => void;
  private _error?: (err: any) => void;

  private _excelFormat: ExcelFormat;
  private _cellFormats: Record<string, CellFormat[]> = {};
  private _content: { sheetName?: string; header?: any; footer?: any } = {};

  constructor(opts: CreateStreamOpts, excelFormat: ExcelFormat) {
    this._workBookWriter = new exceljs.stream.xlsx.WorkbookWriter({ stream: this._stream, useStyles: opts.useStyles });
    this._sheetBegin = opts.sheetBegin;
    this._sheetFinish = opts.sheetFinish;
    this._finish = opts.finish;
    this._error = opts.error;
    this._excelFormat = excelFormat;
  }

  stream(): PassThrough {
    return this._stream;
  }

  add(chunks: any[], sheetIndex?: number) {
    try {
      if (this._workSheet.length === 0) this._workSheet.push(this._workBookWriter.addWorksheet());
      if (!sheetIndex) sheetIndex = 0;

      const workSheet = this._workSheet[sheetIndex];
      chunks.forEach((chunk) => {
        const row = this.createRow(
          chunk,
          workSheet.addRow([]),
          this._cellFormats["TABLE"].filter((e) => e.section === "table" && e.isVariable)
        );
        row.commit();
      });

      this._indexRow++;
    } catch (error) {
      if (this._error) this._error(error);
    }
  }

  setContent(content: { sheetName?: string; header?: any; footer?: any }) {
    this._indexRow = 1;
    this._workSheet.push(this._workBookWriter.addWorksheet(content.sheetName));
    const sheetIndex = this._workSheet.length - 1;

    const cells = this._excelFormat[sheetIndex].cellFomats;

    this._cellFormats = cells.reduce((init, format) => {
      const section = format.section;
      const key = (section ?? "").toUpperCase();

      if (!init[key]) init[key] = [];
      init[key].push(format);

      return init;
    }, {} as Record<string, any[]>);

    if (this._sheetBegin) this._sheetBegin();

    // Add cells in header-section
    this.addHeader(content.header, this._cellFormats["HEADER"], sheetIndex);

    // Add title table
    this.addTitleTable(this._cellFormats["TABLE"], sheetIndex);

    this._content = content;
  }

  async doneSheet(sheetIndex: number) {
    // Add cells in footer-section
    this.addFooter(this._content.footer, this._cellFormats["FOOTER"], sheetIndex, this._indexRow);
    // Set merges and width height
    this.setWidths(this._workSheet[sheetIndex], this._excelFormat[sheetIndex]);
    this._workSheet[sheetIndex].commit();
    if (this._sheetFinish) this._sheetFinish();
  }

  async allDone() {
    await this._workBookWriter.commit();
    if (this._finish) this._finish();
  }

  private createRow(data: any, row: exceljs.Row, cellsFormat: CellFormat[]) {
    if (!data) data = {};

    for (let i = 0; i < cellsFormat.length; i++) {
      const cellFormat = cellsFormat[i];
      const cell = row.getCell(cellFormat.fullAddress.col);
      this.updateCell(data, cellFormat, cell);
    }

    return row;
  }

  private updateCell(rowData: any, cellFormat: CellFormat, cell: exceljs.Cell) {
    cell.value = cellFormat.isVariable ? rowData[(cellFormat.value as any).fieldName] : (cellFormat.value as any).hardValue;
    cell.style = cellFormat.style;
  }

  private addHeader(header: any, cellFormats: CellFormat[], sheetIndex: number) {
    if (!cellFormats) return;

    const workSheet = this._workSheet[sheetIndex];
    const sheetFormat = this._excelFormat[sheetIndex];
    const rowHeaders = [];
    for (let i = 1; i < sheetFormat.beginTableAt; i++) {
      const row = workSheet.addRow([]);
      const formats = cellFormats.filter((e) => e.fullAddress.row === i);
      formats.forEach((format) => this.updateCell(header ?? {}, format, row.getCell(format.fullAddress.col)));

      rowHeaders.push(row);
    }

    rowHeaders.forEach((row) => {
      this.mergesCells(workSheet, sheetFormat, row.number);
    });

    this._indexRow = sheetFormat.beginTableAt;
  }

  private addFooter(footer: any, cellFormats: CellFormat[], sheetIndex: number, endTableAt: number) {
    if (!cellFormats) return;
    const workSheet = this._workSheet[sheetIndex];
    const numberOfRowFooter = cellFormats.reduce((max, val) => (max > val.fullAddress.row ? max : val.fullAddress.row), 0);

    for (let i = 1; i <= numberOfRowFooter; i++) {
      const row = workSheet.addRow([]);
      const formats = cellFormats.filter((e) => this._indexRow === e.fullAddress.row + endTableAt);
      formats.forEach((format) => this.updateCell(footer ?? {}, format, row.getCell(format.fullAddress.col)));
      row.commit();
      this._indexRow++;
    }
  }

  private addTitleTable(cellFormats: CellFormat[], sheetIndex: number) {
    if (!cellFormats) return;
    cellFormats = cellFormats.filter((e) => e.section === "table" && !e.isVariable);
    if (cellFormats.length === 0) return;

    const workSheet = this._workSheet[sheetIndex];
    const numberOfTitleTable = new Set(cellFormats.map((e) => e.fullAddress.row)).size;
    const titleTables = cellFormats.filter((e) => e.section === "table" && !e.isVariable);
    for (let i = 0; i < numberOfTitleTable; i++) {
      const row = workSheet.addRow([]);
      const formats = titleTables.filter((e) => this._indexRow === e.fullAddress.row);
      formats.forEach((format) => this.updateCell({}, format, row.getCell(format.fullAddress.col)));
      row.commit();
      this._indexRow++;
    }
  }

  private mergesCells(sheet: exceljs.Worksheet, sheetFormat: SheetFormat, rowIndex: number) {
    if (sheetFormat.merges) {
      const merges = sheetFormat.merges;
      const keys = Object.keys(merges).filter((key) => merges[key].model.top === rowIndex);
      keys.forEach((masterCell: string) => {
        const { top, left, right, bottom } = merges[masterCell].model;
        sheet.mergeCells(top, left, bottom, right);
      });
    }

    return sheet;
  }

  private setWidths(sheet: exceljs.Worksheet, sheetFormat: SheetFormat) {
    // Set column's width
    sheetFormat.columnWidths?.forEach((colW, i) => {
      if (sheet.columns[i]) sheet.columns[i].width = colW;
    });

    return sheet;
  }
}
