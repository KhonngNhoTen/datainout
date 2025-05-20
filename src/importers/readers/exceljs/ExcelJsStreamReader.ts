import * as exceljs from "exceljs";
import { TypeParser } from "../../../helpers/parse-type.js";
import { Readable } from "stream";
import { TableDataImportHelper } from "../../../helpers/table-data-import-helper.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { CellDataHelper, ExcelReaderHelper } from "../../../helpers/excel-reader-helper.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { SheetSection } from "../../../common/types/common-type.js";

export class ExcelJsStreamReader extends BaseReaderStream {
  private excelReaderHelper: ExcelReaderHelper = new ExcelReaderHelper();
  constructor() {
    super({ type: "excel-stream", typeParser: new TypeParser() });
    this.tableDataImportHelper = new TableDataImportHelper();
  }

  public async load(arg: Readable): Promise<any> {
    this.excelReaderHelper = new ExcelReaderHelper({
      beginTables: this.templates.map((e) => e.beginTableAt),
      endTableAts: this.templates.map((e) => e.endTableAt),
      columnIndexes: this.templates.map((e) => e.keyTableAt),
    });

    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(arg, {
      worksheets: "emit",
    });
    workBookReader.read();

    const that = this;

    (workBookReader as any).on("worksheet", async function (worksheet: exceljs.Worksheet) {
      that.sheetIndex = worksheet.id - 1;
      that.groupCellDescs = that.formatSheet(that.sheetIndex);

      if (that.listEvents.rBegin) that.listEvents.rBegin(worksheet.name);

      (worksheet as any).on("row", function (row: exceljs.Row) {
        if (that.listEvents.rData) that.listEvents.rData();
        that.addRow(row);
      });

      (worksheet as any).on("finished", function () {
        that.addSheet(worksheet);
        if (that.listEvents.rEnd) that.listEvents.rEnd(worksheet.name);
      });
    });

    (workBookReader as any).on("finished", function () {
      if (that.listEvents.rFinish) that.listEvents["rFinish"]();
    });

    return this;
  }

  private addRow(row: exceljs.Row) {
    const cells: CellDataHelper[] = [];
    const sheet = this.templates[this.sheetIndex];
    const section = this.excelReaderHelper.getSection(row.number, this.sheetIndex, (row.values as any[]).slice(1));
    for (let i = 0; i < row.cellCount; i++)
      cells.push(this.excelReaderHelper.getCell(row.getCell(i + 1), row.number, this.sheetIndex, section));

    const trigger = this.tableDataImportHelper.push(cells, sheet, this.groupCellDescs[section], this.chunkSize);

    if (trigger) {
      const filter: FilterImportHandler = {
        section: section,
        sheetIndex: sheet.sheetIndex,
        sheetName: sheet.sheetName,
      };
      const data = this.tableDataImportHelper.pop();
      (async () => await this.callHandlers(data, filter))();
    }
  }

  private addSheet(sheet: exceljs.Worksheet) {
    const filter: FilterImportHandler = {
      section: "footer",
      sheetIndex: sheet.id,
      sheetName: sheet.name,
    };
    const data = this.tableDataImportHelper.pop();
    (async () => await this.callHandlers(data, filter))();
  }
}
