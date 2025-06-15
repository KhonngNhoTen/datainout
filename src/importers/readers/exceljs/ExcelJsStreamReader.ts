import * as exceljs from "exceljs";
import { TypeParser } from "../../../helpers/parse-type.js";
import { Readable } from "stream";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data.js";

export class ExcelJsStreamReader extends BaseReaderStream {
  constructor() {
    super({ type: "excel-stream", typeParser: new TypeParser() });
    this.convertorRows2TableData = new ConvertorRows2TableData();
  }

  public async load(arg: Readable): Promise<any> {
    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(arg, {
      worksheets: "emit",
    });
    workBookReader.read();

    const that = this;

    (workBookReader as any).on("worksheet", async function (worksheet: exceljs.Worksheet) {
      that.sheetIndex = worksheet.id - 1;
      that.groupCellDescs = that.formatSheet(that.sheetIndex);

      // trigger event begin sheet
      that.emitEvent("rBegin", worksheet.name);

      (worksheet as any).on("row", function (row: exceljs.Row) {
        // trigger event load row
        that.emitEvent("rData");
        that.onRow(row);
      });

      (worksheet as any).on("finished", function () {
        // trigger event finish sheet
        that.onSheet(worksheet);
        that.emitEvent("rEnd", worksheet.name);
      });
    });

    // trigger event finish file
    (workBookReader as any).on("finished", function () {
      that.emitEvent("rFinish");
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
      (async () => await this.callHandlers(data, filter))();
    }
  }
}
