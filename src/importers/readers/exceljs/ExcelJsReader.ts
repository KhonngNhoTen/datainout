import { TypeParser } from "../../../helpers/parse-type.js";
import { BaseReader } from "../BaseReader.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";
import { ImporterHandler } from "../../ImporterHandler.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { SheetSection } from "../../../common/types/common-type.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data-v2.js";

export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ReaderExceljsHelper = null as any;

  constructor(templateManager: ExcelTemplateManager<CellImportOptions>) {
    super({ type: "excel", typeParser: new TypeParser(), templateManager });
  }

  async load(arg: unknown): Promise<any> {
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: async () => {
        await this.convertorRows2TableData.push(null);
      },
      onRow: async (data) => {
        await this.convertorRows2TableData.push(data.detail);
      },
      isSampleExcel: false,
      templateManager: this.templateManager,
    });

    this.convertorRows2TableData = new ConvertorRows2TableData({
      onTrigger: async (sect, data) => await this.callHandler(sect, data),
      onErrors: async (err) => await this.onErrors(err),
      chunkSize: this.options?.chunkSize,
      templateManager: this.templateManager,
    });

    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }

  override createTask() {
    return async (section: SheetSection, data: any) => {
      const filter: FilterImportHandler = {
        section: section,
        sheetIndex: this.templateManager.SheetInformation.sheetIndex ?? 0,
        sheetName: this.templateManager.SheetInformation.sheetName,
        isHasNext: data !== null,
      };
      let err;
      data = data instanceof Error ? data : { [section]: data };
      const setGlobalError = (e: any) => (err = e);

      if (this.handler instanceof ImporterHandler) {
        await this.handler.run(data, filter, setGlobalError);
        if (err) throw err;
      } else for (let i = 0; i < this.handler.length; i++) data = await this.handler[i](data, filter);
    };
  }

  private async callHandler(section: SheetSection, data: any) {
    await this.ringPromise.run(section, data);
    // const filter: FilterImportHandler = {
    //   section: section,
    //   sheetIndex: this.templateManager.SheetInformation.sheetIndex ?? 0,
    //   sheetName: this.templateManager.SheetInformation.sheetName,
    //   isHasNext: data !== null,
    // };
    // let err;
    // data = data instanceof Error ? data : { [section]: data };
    // const setGlobalError = (e: any) => (err = e);

    // if (this.handler instanceof ImporterHandler) {
    //   await this.handler.run(data, filter, setGlobalError);
    //   if (err) throw err;
    // } else for (let i = 0; i < this.handler.length; i++) data = await this.handler[i](data, filter);
  }

  override async onErrors(errors: any) {
    errors = Array.isArray(errors) ? errors : [errors];
    if (!this.options?.ignoreErrors) throw errors[0];
    if (this.options?.ignoreErrors) await this.callHandler(null as any, errors[0]);
    else for (let i = 0; i < errors.length; i++) await this.callHandler(null as any, errors[i]);
  }
}
