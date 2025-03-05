import { PassThrough, Stream, Writable } from "stream";
import * as exceljs from "exceljs";
import { ImporterHandler } from "./ImporterHandler.js";
import { ImportFileDesciption, SheetDesciption } from "./ImporterFileDescription.js";
import { CellDescription, FilterImportHandler, SheetSection } from "../type.js";
import { TypeParser } from "../../helper/parse-type.js";
import { getFileExtension } from "../../helper/get-file-extension.js";
import { pathImport } from "../../helper/path-file.js";
import { TableData } from "../../type.js";

type ImporterOptions = {
  importDesciptionPath: string;
  chunkSize?: number;
  handlers: (typeof ImporterHandler)[];
  dateFormat?: string;
};
export class Importer {
  private importDesciption: ImportFileDesciption = [] as any;
  private importDesciptionPath: string;
  private handlers: ImporterHandler[];
  private chunkSize: number;
  private headerTable: CellDescription[] = [];

  private typeParser;

  constructor(opts: ImporterOptions) {
    this.handlers = opts.handlers.map((handler) => new (handler as any)());
    this.importDesciptionPath = opts.importDesciptionPath;
    this.importDesciption = new ImportFileDesciption(require(this.importDesciptionPath));
    this.chunkSize = opts.chunkSize ?? 1;
    this.typeParser = new TypeParser({ dateFormat: opts.dateFormat });
  }

  /**
   * Format cell's value by CellDescription
   */
  private formatValue(cellDescription: CellDescription, value: any, result: TableData) {
    const name = cellDescription.fieldName;
    if (cellDescription.setValue) value = cellDescription.setValue(value, result);
    if (cellDescription.type && cellDescription.type !== "virtual") value = (this.typeParser as any)[cellDescription.type](value);
    if (cellDescription.validate && cellDescription.validate(value)) throw new Error("Validated fail");
    return { name, value };
  }

  /**
   * Get list cell description by section.
   * After, sorting list of cell description.
   * If cell's address is virtual, it is low priority.
   */
  private getCellDescriptionBySection(section: SheetSection, sheetDesciption: SheetDesciption) {
    const cellDescriptions = sheetDesciption.content.filter((e) => e.section === section);
    return cellDescriptions.sort((a, b) => {
      if (!a.address) return 1;
      return a === b ? 0 : a > b ? -1 : 1;
    });
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
   * Read signle section of sheet. Including "Header" and "Footer".
   */
  private async readSingleSectionSheet(section: SheetSection, workSheet: exceljs.Worksheet, sheetDesciption: SheetDesciption) {
    const cellDescriptions = this.getCellDescriptionBySection(section, sheetDesciption);
    const sectionResult = cellDescriptions.reduce((init: undefined | object, val) => {
      const cellValue = val.address ? workSheet.getCell(val.address).value : undefined;
      const format = this.formatValue(val, cellValue, { [section]: init });
      init = { ...(init ?? {}), [format.name]: format.value };
      return init;
    }, undefined);

<<<<<<< Updated upstream:src/imports/reader/Importer.ts
    let result: ResultOfImport = { [section as string]: sectionResult };
    await this.callHandlers(result, { sheetIndex: sheetDesciption.index, section });
=======
    let result: TableData = { [section as string]: sectionResult };
    await this.callHandlers(result, { sheetIndex: sheetDesciption.index, section, sheetName: sheetDesciption.name });
>>>>>>> Stashed changes:src/importers/reader/Importer.ts
  }

  /**
   * Read sheet
   */
  private async readWorkSheet(workBook: exceljs.Workbook, sheetIndex: number) {
    const workSheet = workBook.getWorksheet(sheetIndex + 1);
    const workSheetDescription = this.importDesciption.sheets[sheetIndex];

    if (!workSheet) return;

    // Read header section
    await this.readSingleSectionSheet("header", workSheet, workSheetDescription);

    // Read footer section
    await this.readSingleSectionSheet("footer", workSheet, workSheetDescription);

    // Read table section
    /**
     * Read headers of table.
     * And find cell description by address header.
     * Storing this cell description
     */
    this.headerTable = this.getCellDescriptionBySection("table", workSheetDescription);

    //// Read content of table
<<<<<<< Updated upstream:src/imports/reader/Importer.ts
    const endTable = workSheetDescription.endTable ? workSheetDescription.endTable - 1 : workSheet.rowCount;
    let index = workSheetDescription.startTable + 1;
    let result: ResultOfImport = { table: [] };
=======
    const endTableAt = workSheetDescription.endTableAt
      ? workSheet.rowCount + workSheetDescription.endTableAt
      : workSheet.rowCount;
    let index = workSheetDescription.beginTableAt + 1;
    let result: TableData = { table: [] };
>>>>>>> Stashed changes:src/importers/reader/Importer.ts

    while (index <= endTable) {
      const rows = workSheet.getRows(index, index + this.chunkSize <= endTable ? this.chunkSize : endTable + 1 - index);
      if (rows) {
        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          const values = row.values as any[];
          const resultRow = {} as any;

          if (values.every((e) => e === undefined)) continue;

          this.headerTable.forEach((cellDesc, i) => {
            const format = this.formatValue(cellDesc, values[i + 1], { table: resultRow });
            resultRow[format.name] = format.value;
          });

          result.table?.push(resultRow);
        }
        if (result.table && result.table?.length > 0) await this.callHandlers(result, { sheetIndex, section: "table" });
        result.table = [];
      }

      index += this.chunkSize;
    }
  }

  /**
   * Read file excel and return result after process
   */
  async load(filePath: string): Promise<any>;
  async load(buffer: Buffer): Promise<any>;
  async load(stream: Stream): Promise<any>;
  async load(arg: unknown) {
    const workBook = new exceljs.Workbook();

    if (arg instanceof Buffer) await workBook.xlsx.load(arg);
    else if (typeof arg === "string") await workBook.xlsx.readFile(arg);
    else if (arg instanceof Stream) await workBook.xlsx.read(arg);

    for (let i = 0; i < this.importDesciption.sheets.length; i++) await this.readWorkSheet(workBook, i);
  }

  createStream(): Writable {
    const stream = new PassThrough();

    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(stream, {
      worksheets: "emit",
    });

    workBookReader.read();

    (workBookReader as any).on("worksheet", function (worksheet: any) {
      console.log("worksheet", worksheet);
      worksheet.on("row", function (row: any) {
        console.log(" row.values", row.values);
        console.log(" row.model", row.model);
        console.log("----------");
      });

      worksheet.on("close", function () {
        console.log("worksheet close");
        console.log("----------");
      });

      worksheet.on("finished", function () {
        console.log("worksheet finished");
        console.log("----------");
      });
    });

    (workBookReader as any).on("finished", function () {
      console.log("finished");
      console.log("----------");
    });

    (workBookReader as any).on("close", function () {
      console.log("close");
      console.log("----------");
    });

    return stream;
  }
}
