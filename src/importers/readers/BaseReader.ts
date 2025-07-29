import { ExcelTemplateManager } from "../../common/core/Template.js";
import { CellImportOptions } from "../../common/types/import-template.type.js";
import {
  FilterImportHandler,
  ImporterHandlerInstance,
  ImporterLoadFunctionOpions,
  ImporterReaderType,
} from "../../common/types/importer.type.js";
import { BaseReaderOptions } from "../../common/types/reader.type.js";
import { ConvertorRows2TableData } from "../../helpers/convert-row-to-table-data.js";
import { TypeParser } from "../../helpers/parse-type.js";
import { ImporterHandler } from "../ImporterHandler.js";
import { RingPromise } from "../../common/core/RingPromise.js";
import { QueueData } from "../../common/core/QueueData.js";

export abstract class BaseReader {
  private type: ImporterReaderType;

  protected typeParser: TypeParser;
  protected handler: ImporterHandlerInstance = {} as any;
  protected templateManager: ExcelTemplateManager<CellImportOptions> = {} as any;
  protected convertorRows2TableData: ConvertorRows2TableData = new ConvertorRows2TableData();
  protected options?: ImporterLoadFunctionOpions;
  protected queueData: QueueData<{ data: any | any[]; filter: FilterImportHandler }> = new QueueData(100);
  protected ringPromise: RingPromise = {} as any;
  protected isStopConsumeData: boolean = false;

  constructor(opts: BaseReaderOptions) {
    this.type = opts.type;
    this.typeParser = opts.typeParser ?? new TypeParser();
  }

  protected abstract load(arg: unknown): Promise<any>;

  public async run(
    templateManager: ExcelTemplateManager<CellImportOptions>,
    arg: unknown,
    handler: ImporterHandlerInstance,
    opts?: ImporterLoadFunctionOpions
  ) {
    this.ringPromise = new RingPromise(opts?.jobCount ?? 1, this.createTask());
    this.options = opts;
    this.handler = handler;
    this.templateManager = templateManager;
    this.convertorRows2TableData = new ConvertorRows2TableData({ templateManager });
    await this.load(arg);
    this.consumeData();
  }

  protected async callHandlers(data: any, filter: FilterImportHandler) {
    await this.queueData.waiting();
    this.queueData.add({ data, filter });
  }

  protected async consumeData() {
    while (true && !this.isStopConsumeData) {
      const queueData = this.queueData.shift();
      if (queueData && queueData.data) {
        await this.ringPromise.run(queueData?.data, queueData?.filter);
      } else if (!queueData?.data) break;
      else if (!queueData) await new Promise((r) => setTimeout(r, 2));
    }
  }

  public getType() {
    return this.type;
  }

  protected createTask() {
    return async (data: any | any[], filter: FilterImportHandler) => {
      if (this.handler instanceof ImporterHandler) {
        return await this.handler.run(data, filter);
      } else {
        for (let i = 0; i < this.handler.length; i++) {
          data = await this.handler[i](data, filter);
        }
      }
    };
  }
}
