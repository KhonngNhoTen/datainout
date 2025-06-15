import * as exceljs from "exceljs";
import { TypeParser } from "../../../helpers/parse-type.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { RowDataHelper, SheetDataHelper } from "../../../common/types/excel-reader-helper.type.js";

export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ReaderExceljsHelper = new ReaderExceljsHelper();

  constructor() {
    super({ type: "excel", typeParser: new TypeParser() });
  }

  async load(arg: unknown): Promise<any> {
    const _that = this;
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: _that.onSheet,
      onRow: _that.onRow,
      isSampleExcel: false,
      template: this.templates,
    });

    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }

  private async onRow(row: RowDataHelper) {
    const sheet = this.templates[this.sheetIndex];
    this.handleRow({ id: sheet.sheetIndex + 1, name: sheet.sheetName } as any, row.detail);
  }

  private async onSheet(sheet: SheetDataHelper) {
    this.handleRow(sheet.detail, null);
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
