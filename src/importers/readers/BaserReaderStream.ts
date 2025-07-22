import { EventRegister, IBaseStream } from "../../common/core/ListEvents.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { EventType } from "../../common/types/common-type.js";
import { CellImportOptions, SheetImportOptions } from "../../common/types/import-template.type.js";
import { ImporterHandlerInstance } from "../../common/types/importer.type.js";
import { TypeParser } from "../../helpers/parse-type.js";
import { BaseReader } from "./BaseReader.js";
import { Readable } from "stream";

export abstract class BaseReaderStream extends BaseReader implements IBaseStream {
  protected listEvents: EventRegister = new EventRegister();
  protected readable: Readable;
  protected handler: ImporterHandlerInstance;

  constructor(templateManager: ExcelTemplateManager<CellImportOptions>, readable: Readable, handler: ImporterHandlerInstance) {
    super({ type: "excel-stream", typeParser: new TypeParser() });
    this.templateManager = templateManager;
    this.handler = handler;
    this.readable = readable;
  }
  onStart(func: EventType["start"]): this {
    throw new Error("Method not implemented.");
  }
  endData(func: EventType["enddata"]): this {
    throw new Error("Method not implemented.");
  }
  onHeader(func: EventType["header"]): this {
    throw new Error("Method not implemented.");
  }
  onFooter(func: EventType["footer"]): this {
    throw new Error("Method not implemented.");
  }

  start(): void {
    (async () => await this.run(this.templateManager, this.readable, this.handler))();
  }

  on<EventKey extends keyof EventType>(key: EventKey, func: EventType[EventKey]): this {
    this.listEvents.on(key, func);
    return this;
  }

  onFinish(func: EventType["finish"]): this {
    this.listEvents.onFinish(func);
    return this;
  }

  onBegin(func: EventType["begin"]): this {
    this.listEvents.onBegin(func);
    return this;
  }

  onData(func: EventType["data"]): this {
    this.listEvents.onData(func);
    return this;
  }

  onEnd(func: EventType["end"]): this {
    this.listEvents.onEnd(func);
    return this;
  }

  onError(func: EventType["error"]): this {
    this.listEvents.onError(func);
    return this;
  }

  protected abstract override load(arg: unknown): Promise<any>;
}
