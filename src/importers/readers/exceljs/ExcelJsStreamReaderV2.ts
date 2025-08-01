import * as exceljs from "exceljs";
import { Readable } from "stream";
import { ImporterHandlerInstance } from "../../../common/types/importer.type.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";

export class ExcelJsStreamReader extends BaseReaderStream {
  constructor(templateManager: ExcelTemplateManager<CellImportOptions>, readable: Readable, handlers: ImporterHandlerInstance) {
    super(templateManager, readable, handlers);
  }

  public async load(arg: Readable): Promise<any> {
    const workBookReader = new exceljs.stream.xlsx.WorkbookReader(arg, {
      worksheets: "emit",
    });
    // workBookReader.read();
    for await (const worksheet of workBookReader) {
      const sheet = worksheet as unknown as exceljs.Worksheet;
      this.listEvents.emitEvent("begin", sheet.name);
      for await (const row of worksheet) {
        // Check exist glabal Error, then throw Error;
        if (this.globalError) throw this.globalError;

        this.templateManager.defineActualTableStartRow(ReaderExceljsHelper.beginTableAt(row, this.templateManager.SheetTemplate, false));
        this.templateManager.defineActualTableEndRow(ReaderExceljsHelper.endTableAt(row, this.templateManager.SheetTemplate, false));
        await this.convertorRows2TableData.push(row);
        this.listEvents.emitEvent("data");
      }
      this.listEvents.emitEvent("end", sheet.name);
    }
    await this.convertorRows2TableData.push(null);
    this.listEvents.emitEvent("finish");
    return this;
  }
}
