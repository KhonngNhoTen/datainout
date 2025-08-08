import { ExcelTemplateManager } from "../../common/core/Template.js";
import { CellImportOptions } from "../../common/types/import-template.type.js";
import {
  FilterImportHandler,
  ImporterHandlerInstance,
  ImporterLoadFunctionOpions,
  ImporterReaderType,
} from "../../common/types/importer.type.js";
import { BaseReaderOptions } from "../../common/types/reader.type.js";
import { ConvertorRows2TableData } from "../../helpers/convert-row-to-table-data-v2.js";
import { TypeParser } from "../../helpers/parse-type.js";
import { ImporterHandler } from "../ImporterHandler.js";
import { RingPromise } from "../../common/core/RingPromise.js";
import { QueueData } from "../../common/core/QueueData.js";
import { SheetSection } from "../../common/types/common-type.js";

export abstract class BaseReader {
  private type: ImporterReaderType;

  protected typeParser: TypeParser;
  protected handler: ImporterHandlerInstance = {} as any;
  protected templateManager: ExcelTemplateManager<CellImportOptions>;
  protected convertorRows2TableData: ConvertorRows2TableData;
  protected options?: ImporterLoadFunctionOpions;
  protected queueData: QueueData<{ data: any | any[]; filter: FilterImportHandler }> = new QueueData(100);
  protected ringPromise: RingPromise = {} as any;
  protected globalError?: Error;

  constructor(opts: BaseReaderOptions) {
    this.type = opts.type;
    this.typeParser = opts.typeParser ?? new TypeParser();
    this.templateManager = opts.templateManager;
    this.convertorRows2TableData = new ConvertorRows2TableData({
      onTrigger: async (sect, data) => await this.callHandler(sect, data),
      onErrors: async (err) => await this.onErrors(err),
      chunkSize: opts?.chunkSize,
      templateManager: this.templateManager,
    });
  }

  protected abstract load(arg: unknown): Promise<any>;

  public async run(
    templateManager: ExcelTemplateManager<CellImportOptions>,
    arg: unknown,
    handler: ImporterHandlerInstance,
    opts?: ImporterLoadFunctionOpions
  ) {
    this.convertorRows2TableData = new ConvertorRows2TableData({
      onTrigger: async (sect, data) => await this.callHandler(sect, data),
      onErrors: async (err) => await this.onErrors(err),
      chunkSize: opts?.chunkSize,
      templateManager: this.templateManager,
    });
    this.ringPromise = new RingPromise(opts?.jobCount ?? 1, this.createTask());
    this.options = opts;
    this.handler = handler;
    this.templateManager = templateManager;

    await this.load(arg);
  }

  public getType() {
    return this.type;
  }

  protected setGlobalError(err: Error) {
    if (!this.options?.ignoreErrors && this.globalError) this.globalError = err;
  }

  protected createTask() {
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

  protected async callHandler(section: SheetSection, data: any) {
    await this.ringPromise.run(section, data);
  }

  protected async onErrors(errors: any) {
    errors = Array.isArray(errors) ? errors : [errors];
    if (!this.options?.ignoreErrors) throw errors[0];
    if (this.options?.ignoreErrors) await this.callHandler(null as any, errors[0]);
    else for (let i = 0; i < errors.length; i++) await this.callHandler(null as any, errors[i]);
  }
}
