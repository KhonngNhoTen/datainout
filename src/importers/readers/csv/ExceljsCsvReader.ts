import * as exceljs from "exceljs";
import { Readable } from "stream";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";
import { SheetSection } from "../../../common/types/common-type.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data.js";

export class ExcelJsCsvReader extends BaseReader {
  constructor() {
    super({ type: "csv" });
  }
  public async load(arg: unknown): Promise<any> {
    this.convertorRows2TableData = new ConvertorRows2TableData({
      chunkSize: this.options?.chunkSize,
      templateManager: this.templateManager,
    });
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
      await this.handleRow(workSheet, row);
    }

    // End file

    await this.handleRow(workSheet, null);
  }

  private async handleRow(workSheet: exceljs.Worksheet, section: SheetSection, rowIndex: number): Promise<void>;
  private async handleRow(workSheet: exceljs.Worksheet, row: exceljs.Row | null): Promise<void>;
  private async handleRow(workSheet: exceljs.Worksheet, arg: unknown, rowIndex?: number) {
    const { isTrigger, triggerSection, hasError, errors } = rowIndex
      ? this.convertorRows2TableData.pushBySection(arg as SheetSection, this.templateManager.SheetTemplate, rowIndex)
      : this.convertorRows2TableData.push(arg as any, this.templateManager.SheetTemplate);

    if (hasError) await this.handleError(errors);

    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: triggerSection,
        sheetIndex: workSheet.id - 1,
        sheetName: workSheet.name,
        isHasNext: arg !== null,
      };
      const data = this.convertorRows2TableData.pop(triggerSection);
      await this.callHandlers(data, filter);
    }
  }

  private async handleError(errors: Error[]) {
    if (errors.length === 0) return;
    if (this.options?.ignoreErrors === true) throw errors[0];

    for (let i = 0; i < errors.length; i++) {
      await this.callHandlers(errors[i], null as any);
    }
  }
}
