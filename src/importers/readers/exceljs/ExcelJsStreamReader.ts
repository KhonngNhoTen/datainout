import * as exceljs from "exceljs";
import { TypeParser } from "../../../helpers/parse-type.js";
import { Readable } from "stream";
import { TableDataImportHelper } from "../../../helpers/table-data-import-helper.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { CellDataHelper, ExcelReaderHelper } from "../../../helpers/excel-reader-helper.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { IReader } from "../../../common/decorators/IReader.decorator.js";

@IReader()
export class ExcelJsStreamReader extends BaseReaderStream {
  private excelReaderHelper: ExcelReaderHelper = new ExcelReaderHelper();
  constructor() {
    super({ type: "excel-stream", typeParser: new TypeParser() });
    this.tableDataImportHelper = new TableDataImportHelper();
  }

  public async load(arg: Readable): Promise<any> {
    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(arg, {
      worksheets: "emit",
    });
    workBookReader.read();

    const that = this;

    (workBookReader as any).on("worksheet", async function (worksheet: exceljs.Worksheet) {
      that.sheetIndex = worksheet.id;
      that.groupCellDescs = that.formatSheet(that.sheetIndex);

      if (that.listEvents.rBegin) that.listEvents.rBegin(worksheet.name);

      (async () => await that.addSheet(worksheet))();

      (worksheet as any).on("row", function (row: exceljs.Row) {
        if (that.listEvents.rData) that.listEvents.rData();
        (async () => await that.addRow(row))();
      });

      (worksheet as any).on("finished", function () {
        if (that.listEvents.rEnd) that.listEvents.rEnd(worksheet.name);
      });
    });

    (workBookReader as any).on("finished", function () {
      if (that.listEvents.rFinish) that.listEvents["rFinish"]();
    });

    return this;
  }

  private async addCell(cell: exceljs.Cell) {
    return this.excelReaderHelper.getCell(cell, cell.fullAddress.row, this.sheetIndex);
  }

  private async addRow(row: exceljs.Row) {
    const cells: CellDataHelper[] = [];
    const sheet = this.templates[this.sheetIndex];
    const rowData = this.excelReaderHelper.getRow(row, sheet.beginTableAt, sheet.endTableAt);
    for (let i = 0; i < row.cellCount; i++) cells.push(await this.addCell(row.getCell(i + 1)));

    const isTrigger = this.tableDataImportHelper.push(cells, sheet, this.groupCellDescs[cells[0].section], this.chunkSize);

    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: rowData.section,
        sheetIndex: rowData.rowIndex,
        sheetName: sheet.sheetName,
      };
      const data = this.tableDataImportHelper.get();
      await this.callHandlers(data, filter);
    }
  }

  private async addSheet(sheet: exceljs.Worksheet) {
    const filter: FilterImportHandler = {
      section: "footer",
      sheetIndex: this.sheetIndex,
      sheetName: sheet.name,
    };
    const data = this.tableDataImportHelper.get();
    await this.callHandlers(data, filter);
  }
}
