import * as exceljs from "exceljs";
import { Readable } from "stream";
import { ImporterBaseReaderStreamType, ImporterHandlerInstance, ImporterLoadFunctionOpions } from "../../../common/types/importer.type.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";

export class ExcelJsStreamReader extends BaseReaderStream {
  constructor(data: {
    templateManager: ExcelTemplateManager<CellImportOptions>;
    readable: Readable;
    handler: ImporterHandlerInstance;
    options?: ImporterLoadFunctionOpions & { type?: ImporterBaseReaderStreamType };
  }) {
    super(data);
  }

  public async load(arg: Readable): Promise<any> {
    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(arg, {
      worksheets: "emit",
    });
    workBookReader.read();

    const that = this;

    (workBookReader as any).on("worksheet", async function (worksheet: exceljs.Worksheet) {
      that.templateManager.SheetIndex = worksheet.id - 1;
      // trigger event begin sheet
      that.listEvents.emitEvent("begin", worksheet.name);

      (worksheet as any).on("row", function (row: exceljs.Row) {
        // trigger event load row
        that.templateManager.defineActualTableStartRow(ReaderExceljsHelper.beginTableAt(row, that.templateManager.SheetTemplate, false));
        that.templateManager.defineActualTableEndRow(ReaderExceljsHelper.endTableAt(row, that.templateManager.SheetTemplate, false));
        that.convertorRows2TableData.push(row);
        that.listEvents.emitEvent("data");
      });

      (worksheet as any).on("finished", function () {
        // trigger event finish sheet
        that.convertorRows2TableData.push(null);
        that.listEvents.emitEvent("end", worksheet.name);
      });
    });

    // trigger event finish file
    (workBookReader as any).on("finished", function () {
      that.listEvents.emitEvent("finish");
    });

    return this;
  }
}
