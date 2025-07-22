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
import { ConvertorRows2TableData2 } from "../../helpers/convert-row-to-table-data-v2.js";

export abstract class BaseReader {
  private type: ImporterReaderType;

  protected typeParser: TypeParser;
  protected handler: ImporterHandlerInstance = {} as any;
  protected templateManager: ExcelTemplateManager<CellImportOptions> = {} as any;
  protected convertorRows2TableData: ConvertorRows2TableData = new ConvertorRows2TableData();

  protected batches: { func: (...arg: any[]) => Promise<void>; params: any }[] = [];
  protected options?: ImporterLoadFunctionOpions;
  protected BATCH_MAX_SIZE = 2;
  protected ringPromise: RingPromise = {} as any;

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
    this.options = opts;
    this.handler = handler;
    this.templateManager = templateManager;
    this.ringPromise = new RingPromise(opts?.workerSize ?? 1, this.createTask());
    await this.load(arg);
  }

  protected async callHandlers(data: any, filter: FilterImportHandler) {
    await this.ringPromise.run(data, filter);
  }

  protected createTask() {
    return async (data: any, filter: FilterImportHandler) => {
      if (this.handler) {
        if (this.handler instanceof ImporterHandler) {
          return await this.handler.run(data, filter);
        } else {
          for (let i = 0; i < this.handler.length; i++) {
            data = await this.handler[i](data, filter);
          }
        }
      }
    };
  }

  public getType() {
    return this.type;
  }
}
