import * as exceljs from "exceljs";
import { Readable } from "stream";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";

export class ExcelJsCsvReader extends BaseReader {
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
      const row = workSheet.getRow(i + 1);
      this.handleRow(workSheet, row);
    }

    // End file
    this.handleRow(workSheet, null);
  }

  private handleRow(workSheet: exceljs.Worksheet, row: exceljs.Row | null) {
    const { isTrigger, triggerSection } = this.convertorRows2TableData.push(row, this.templates[this.sheetIndex]);

    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: triggerSection,
        sheetIndex: workSheet.id - 1,
        sheetName: workSheet.name,
      };
      const data = this.convertorRows2TableData.pop(triggerSection);
      (async () => await this.callHandlers(data, filter))();
    }
  }
}
