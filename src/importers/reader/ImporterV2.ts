import { PassThrough, Stream, Writable } from "stream";
import * as exceljs from "exceljs";
import { ImporterHandler } from "./ImporterHandler.js";
import { ImportFileDesciption, SheetDesciption } from "./ImporterFileDescription.js";
import { CellDescription, FilterImportHandler, SheetSection } from "../type.js";
import { TypeParser } from "../../helper/parse-type.js";
import { getFileExtension } from "../../helper/get-file-extension.js";
import { pathImport } from "../../helper/path-file.js";
import { TableData } from "../../type.js";
import { ChunkData, TableDataHandler } from "../../common/TableDataHandler.js";
import { sortByAddress } from "../../helper/sort-by-address.js";

type ImporterOptions = {
  templatePath: string;
  handlers: (typeof ImporterHandler)[];
  dateFormat?: string;
};
export class Importer {
  private importDesciption: ImportFileDesciption = [] as any;
  private templatePath: string;
  private handlers: ImporterHandler[];
  private typeParser;

  constructor(opts: ImporterOptions) {
    this.handlers = opts.handlers.map((handler) => new (handler as any)());
    this.typeParser = new TypeParser({ dateFormat: opts.dateFormat });
    this.templatePath = pathImport(opts.templatePath, "templateDir");
    this.importDesciption =
      getFileExtension(this.templatePath) === "js"
        ? new ImportFileDesciption(require(this.templatePath))
        : new ImportFileDesciption(require(this.templatePath).default);
  }

  /**
   * Format cell's value by CellDescription
   */
  private formatValue(cellDescription: CellDescription, value: any, row: {}) {
    if (cellDescription.setValue) value = cellDescription.setValue(value, row);
    if (cellDescription.type && cellDescription.type !== "virtual") value = (this.typeParser as any)[cellDescription.type](value);
    if (cellDescription.validate && cellDescription.validate(value)) throw new Error("Validated fail");
    return value;
  }

  /**
   * Call list of handler with argument is result
   */
  private async callHandlers(result: TableData, filter: FilterImportHandler) {
    for (let i = 0; i < this.handlers.length; i++) {
      result = await this.handlers[i].run(result, filter);
    }
  }

  /**
   * Create Table data from ChunkData
   */
  private async convertTableData(data: ChunkData, filter: FilterImportHandler, sheetDesciption: SheetDesciption) {
    const section = filter.section;
    const cells: exceljs.Cell[][] = data.table && data.table.length > 0 ? data.table : ([data[section]] as exceljs.Cell[][]);
    const cellDescs = sortByAddress(sheetDesciption.content.filter((e) => e.section === filter.section));

    const tableData: TableData = { [section]: section === "table" ? [] : {} };
    let content: any = {};
    for (let i = 0; i < cells.length; i++) {
      content = {};
      const sortedCells = sortByAddress(cells[i]);
      const rawValues: any = sortedCells.reduce((init, e, j) => ({ ...init, [cellDescs[j].fieldName]: e.value }), {});
      cellDescs.forEach((e, j) => (content[e.fieldName] = this.formatValue(e, rawValues[e.fieldName], rawValues)));
      if (section === "table") tableData.table?.push(content);
      else tableData[section] = content;
    }

    await this.callHandlers(tableData, filter);
  }

  /**
   * Read sheet
   */
  private async readWorkSheet(workSheet: exceljs.Worksheet, chunkSize?: number) {
    const sheetDesciption: SheetDesciption = this.importDesciption.sheets[workSheet.id - 1];
    const tableDataHandler = new TableDataHandler({
      beginTableAt: sheetDesciption.beginTableAt,
      keyIndex: sheetDesciption.keyIndex,
      chunkSize,
      sheetIndex: workSheet.id,
      sheetName: workSheet.name,
      sheetDesciption,
    });

    tableDataHandler.on("table", (data, filter) => (async () => await this.convertTableData(data, filter, sheetDesciption))());
    tableDataHandler.on("header", (data, filter) => (async () => await this.convertTableData(data, filter, sheetDesciption))());
    tableDataHandler.on("footer", (data, filter) => (async () => await this.convertTableData(data, filter, sheetDesciption))());

    for (let i = 1; i <= workSheet.rowCount; i++) tableDataHandler.addRow(workSheet.getRow(i));
    tableDataHandler.addRow(null);
  }

  /**
   * Read file excel and return result after process
   */
  async load(filePath: string, chunkSize?: number): Promise<any>;
  async load(buffer: Buffer, chunkSize?: number): Promise<any>;
  async load(stream: Stream, chunkSize?: number): Promise<any>;
  async load(arg: unknown, chunkSize?: number) {
    const workBook = new exceljs.Workbook();

    if (arg instanceof Buffer) await workBook.xlsx.load(arg as unknown as exceljs.Buffer);
    else if (typeof arg === "string") await workBook.xlsx.readFile(arg);
    else if (arg instanceof Stream) await workBook.xlsx.read(arg);

    for (let i = 0; i < this.importDesciption.sheets.length; i++) {
      const workSheet = workBook.getWorksheet(i + 1);
      if (workSheet) await this.readWorkSheet(workSheet, chunkSize);
    }
  }

  createStream(opts?: {
    sheetFinished?: () => void;
    workBookFinished?: () => void;
    sheetBegin?: (worksheet: exceljs.Worksheet, sheetDesciption: SheetDesciption) => void;
    chunkSize?: number;
  }): Writable {
    const stream = new PassThrough();

    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(stream, {
      worksheets: "emit",
    });
    workBookReader.read();

    let tableDataHandler: TableDataHandler;
    const that = this;

    (workBookReader as any).on("worksheet", async function (worksheet: exceljs.Worksheet) {
      const sheetDesciption = that.importDesciption.sheets[worksheet.id - 1];
      if (!tableDataHandler) {
        tableDataHandler = new TableDataHandler({
          beginTableAt: sheetDesciption.beginTableAt,
          keyIndex: sheetDesciption.keyIndex,
          chunkSize: opts?.chunkSize ?? 10,
          sheetIndex: worksheet.id,
          sheetName: worksheet.name,
          sheetDesciption,
        });
        tableDataHandler.on("table", (data, filter) =>
          (async () => await that.convertTableData(data, filter, sheetDesciption))()
        );
        tableDataHandler.on("header", (data, filter) =>
          (async () => await that.convertTableData(data, filter, sheetDesciption))()
        );
        tableDataHandler.on("footer", (data, filter) =>
          (async () => await that.convertTableData(data, filter, sheetDesciption))()
        );
      }

      (worksheet as any).on("row", function (row: exceljs.Row) {
        tableDataHandler.addRow(row);
      });
      if (opts?.sheetBegin) opts.sheetBegin(worksheet, sheetDesciption);

      (worksheet as any).on("finished", function () {
        tableDataHandler.addRow(null);
        if (opts?.sheetFinished) opts.sheetFinished();
      });
    });

    (workBookReader as any).on("finished", function () {
      if (opts?.workBookFinished) opts.workBookFinished();
    });

    return stream;
  }
}
