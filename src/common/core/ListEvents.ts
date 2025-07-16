import { EventType } from "../types/common-type.js";
export interface IEventImporterRegister {
  on<EventKey extends keyof EventType>(key: EventKey, func: EventType[EventKey]): this;
  onFinish(func: EventType["finish"]): this;
  onBegin(func: EventType["begin"]): this;
  onData(func: EventType["data"]): this;
  endData(func: EventType["enddata"]): this;
  onEnd(func: EventType["end"]): this;
  onError(func: EventType["error"]): this;
  onHeader(func: EventType["header"]): this;
  onFooter(func: EventType["footer"]): this;
}

export interface IBaseStream extends IEventImporterRegister {
  /** Start stream */
  start(): void;
}

export class EventRegister implements IEventImporterRegister {
  private listEvents: Partial<EventType> = {};

  on<EventKey extends keyof EventType>(key: EventKey, func: EventType[EventKey]): this {
    this.listEvents[key] = func;
    return this;
  }

  onFinish(func: EventType["finish"]): this {
    this.listEvents.finish = func;
    return this;
  }
  onFile(func: EventType["onFile"]): this {
    this.listEvents.onFile = func;
    return this;
  }

  onBegin(func: EventType["begin"]): this {
    this.listEvents.begin = func;
    return this;
  }

  onData(func: EventType["data"]): this {
    this.listEvents.data = func;
    return this;
  }

  endData(func: EventType["enddata"]): this {
    this.listEvents.enddata = func;
    return this;
  }

  onEnd(func: EventType["end"]): this {
    this.listEvents.end = func;
    return this;
  }

  onError(func: EventType["error"]): this {
    this.listEvents.error = func;
    return this;
  }

  onHeader(func: EventType["header"]): this {
    this.listEvents.header = func;
    return this;
  }
  onFooter(func: EventType["footer"]): this {
    this.listEvents.footer = func;
    return this;
  }

  public emitEvent(key: keyof EventType, data?: any) {
    if (this.listEvents[key]) this.listEvents[key](data);
    return this;
  }
}
