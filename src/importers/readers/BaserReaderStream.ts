import { EventRegister, IBaseStream, IEventRegister } from "../../common/core/ListEvents.js";
import { EventType } from "../../common/types/common-type.js";
import { SheetImportOptions } from "../../common/types/import-template.type.js";
import { ImporterBaseReaderStreamType, ImporterHandlerFunction } from "../../common/types/importer.type.js";
import { TypeParser } from "../../helpers/parse-type.js";
import { BaseReader } from "./BaseReader.js";
import { Readable } from "stream";

export abstract class BaseReaderStream extends BaseReader implements IBaseStream {
  protected listEvents: EventRegister = new EventRegister();
  protected templates: SheetImportOptions[];
  protected readable: Readable;
  protected handlers: ImporterHandlerFunction[];

  constructor(templates: SheetImportOptions[], readable: Readable, handlers: ImporterHandlerFunction[]) {
    super({ type: "excel-stream", typeParser: new TypeParser() });
    this.templates = templates;
    this.handlers = handlers;
    this.readable = readable;
  }

  start(): void {
    (async () => await this.run(this.templates, this.readable, this.handlers))();
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
