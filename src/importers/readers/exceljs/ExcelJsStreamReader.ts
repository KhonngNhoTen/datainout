import * as exceljs from "exceljs";
import { Readable } from "stream";
import { FilterImportHandler, ImporterHandlerInstance } from "../../../common/types/importer.type.js";
import { BaseReaderStream } from "../BaserReaderStream.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";
import { SheetSection } from "../../../common/types/common-type.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";

export class ExcelJsStreamReader extends BaseReaderStream {
  constructor(templateManager: ExcelTemplateManager<CellImportOptions>, readable: Readable, handlers: ImporterHandlerInstance) {
    super(templateManager, readable, handlers);
  }

  public async load(arg: Readable): Promise<any> {
    this.convertorRows2TableData = new ConvertorRows2TableData({
      chunkSize: this.options?.chunkSize,
      templateManager: this.templateManager,
    });
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
    const sheet = this.templateManager.SheetTemplate;
    this.handleRow({ id: sheet.sheetIndex + 1, name: sheet.sheetName } as any, row);
  }

  private onSheet(sheet: exceljs.Worksheet) {
    this.handleRow(sheet, null);
  }

  private handleRow(workSheet: exceljs.Worksheet, section: SheetSection, rowIndex: number): void;
  private handleRow(workSheet: exceljs.Worksheet, row: exceljs.Row | null): void;
  private handleRow(workSheet: exceljs.Worksheet, arg: unknown, rowIndex?: number) {
    const { isTrigger, triggerSection, hasError, errors } = rowIndex
      ? this.convertorRows2TableData.pushBySection(arg as SheetSection, this.templateManager.SheetTemplate, rowIndex)
      : this.convertorRows2TableData.push(arg as any, this.templateManager.SheetTemplate);

    if (hasError) this.handleError(errors);
    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: triggerSection,
        sheetIndex: workSheet.id - 1,
        sheetName: workSheet.name,
        isHasNext: arg !== null,
      };
      const data = this.convertorRows2TableData.pop(triggerSection);
      (async () => await this.callHandlers(data, filter))();
    }
  }

  private handleError(errors: Error[]) {
    if (errors.length === 0) return;
    if (this.options?.ignoreErrors === true) throw errors[0];

    for (let i = 0; i < errors.length; i++) {
      (async () => await this.callHandlers(errors[i], null as any))();
    }
  }
}
