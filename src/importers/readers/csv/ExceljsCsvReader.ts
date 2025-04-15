import * as exceljs from "exceljs";
import { Readable } from "stream";
import { CellDataHelper, ExcelReaderHelper } from "../../../helpers/excel-reader-helper.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";
import { IReader } from "../../../common/decorators/IReader.decorator.js";

@IReader()
export class ExcelJsCsvReader extends BaseReader {
  private excelReaderHelper: ExcelReaderHelper = new ExcelReaderHelper();

  constructor() {
    super({ type: "csv" });
  }
  public async load(arg: unknown): Promise<any> {
    let workSheet: exceljs.Worksheet;
    const workBook = new exceljs.Workbook();
    if (arg instanceof Buffer) {
      const stream = new Readable();
      workSheet = await workBook.csv.read(stream);
      stream.push(arg);
      stream.push(null);
    } else workSheet = await workBook.csv.readFile(arg as string);

    for (let i = 0; i < workSheet.actualRowCount; i++) {
      await this.readRow(workSheet.getRow(i + 1));
    }

    // End file
    const data = this.tableDataImportHelper.get();
    await this.callHandlers(data, {
      section: "footer",
      sheetIndex: 0,
      sheetName: "",
    });
  }

  private async readRow(row: exceljs.Row) {
    const cells: CellDataHelper[] = [];
    for (let i = 0; i < row.cellCount; i++) {
      cells.push(this.excelReaderHelper.getCell(row.getCell(i + 1), row.number, 0));
    }

    const isTrigger = this.tableDataImportHelper.push(
      cells,
      this.templates[this.sheetIndex],
      this.groupCellDescs[cells[0].section],
      this.chunkSize
    );

    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: cells[0].section,
        sheetIndex: cells[0].rowIndex,
        sheetName: this.templates[this.sheetIndex].sheetName,
      };
      const data = this.tableDataImportHelper.get();
      await this.callHandlers(data, filter);
    }
  }
}
