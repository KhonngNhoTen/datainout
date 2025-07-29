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
    // workBookReader.read();
    for await (const worksheet of workBookReader) {
      const sheet = worksheet as unknown as exceljs.Worksheet;
      this.listEvents.emitEvent("begin", sheet.name);
      this.onSheet(sheet);
      for await (const row of worksheet) {
        this.onRow(row);
        this.listEvents.emitEvent("data");
      }
      this.listEvents.emitEvent("end", sheet.name);
    }
    this.listEvents.emitEvent("finish");
    return this;
  }

  private async onRow(row: exceljs.Row) {
    const sheet = this.templateManager.SheetTemplate;
    await this.handleRow({ id: sheet.sheetIndex + 1, name: sheet.sheetName } as any, row);
  }

  private async onSheet(sheet: exceljs.Worksheet) {
    // await this.handleRow(sheet, null);
    await this.callHandlers(null, null as any);
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
      await this.callHandlers({ [triggerSection]: data }, filter);
    }
  }

  private async handleError(errors: Error[]) {
    if (errors.length === 0) return;
    if (this.options?.ignoreErrors === true) {
      this.isStopConsumeData = true;
      throw errors[0];
    }

    for (let i = 0; i < errors.length; i++) {
      await this.callHandlers(errors[i], null as any);
    }
  }
}
