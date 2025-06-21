import * as exceljs from "exceljs";
import { Readable } from "stream";
import { FilterImportHandler, ImporterHandlerFunction } from "../../../common/types/importer.type.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data.js";

export class ExcelJsStreamReader extends BaseReaderStream {
  constructor(templatePath: string, readable: Readable, handlers: ImporterHandlerFunction[]) {
    super(templatePath, readable, handlers);
  }

  public async load(arg: Readable): Promise<any> {
    this.convertorRows2TableData = new ConvertorRows2TableData();
    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(arg, {
      worksheets: "emit",
    });
    workBookReader.read();

    const that = this;

    (workBookReader as any).on("worksheet", async function (worksheet: exceljs.Worksheet) {
      that.sheetIndex = worksheet.id - 1;
      that.groupCellDescs = that.formatSheet(that.sheetIndex);

      // trigger event begin sheet
      that.listEvents.emitEvent("begin", worksheet.name);

      (worksheet as any).on("row", function (row: exceljs.Row) {
        // trigger event load row
        that.onRow(row);
        that.listEvents.emitEvent("data");
      });

      (worksheet as any).on("finished", function () {
        // trigger event finish sheet
        that.onSheet(worksheet);
        that.listEvents.emitEvent("end", worksheet.name);
      });
    });

    // trigger event finish file
    (workBookReader as any).on("finished", function () {
      that.listEvents.emitEvent("finish");
    });

    return this;
  }

  private onRow(row: exceljs.Row) {
    const sheet = this.templates[this.sheetIndex];
    this.handleRow({ id: sheet.sheetIndex + 1, name: sheet.sheetName } as any, row);
  }

  private onSheet(sheet: exceljs.Worksheet) {
    this.handleRow(sheet, null);
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
      (async () => await this.callHandlers({ [triggerSection]: data }, filter))();
    }
  }
}
