import * as exceljs from "exceljs";
import { Readable } from "stream";
import { BaseReader } from "../BaseReader.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";

export class ExcelJsCsvReader extends BaseReader {
  constructor(templateManager: ExcelTemplateManager<CellImportOptions>) {
    super({ type: "csv", templateManager });
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
      this.templateManager.defineActualTableStartRow(ReaderExceljsHelper.beginTableAt(row, this.templateManager.SheetTemplate, false));
      this.templateManager.defineActualTableEndRow(ReaderExceljsHelper.endTableAt(row, this.templateManager.SheetTemplate, false));
      await this.convertorRows2TableData.push(row);
    }

    // End file
    await this.convertorRows2TableData.push(null);
  }
}
