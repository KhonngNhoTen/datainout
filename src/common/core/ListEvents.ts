import { EventType } from "../types/common-type.js";
export interface IEventRegister {
  on<EventKey extends keyof EventType>(key: EventKey, func: EventType[EventKey]): this;
  onFinish(func: EventType["finish"]): this;
  onBegin(func: EventType["begin"]): this;
  onData(func: EventType["data"]): this;
  onEnd(func: EventType["end"]): this;

  onError(func: EventType["error"]): this;
}

export interface IBaseStream extends IEventRegister {
  /** Start stream */
  start(): void;
}

export class EventRegister implements IEventRegister {
  private listEvents: Partial<EventType> = {};

  on<EventKey extends keyof EventType>(key: EventKey, func: EventType[EventKey]): this {
    this.listEvents[key] = func;
    return this;
  }

  onFinish(func: EventType["finish"]): this {
    this.listEvents.finish = func;
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

  onEnd(func: EventType["end"]): this {
    this.listEvents.end = func;
    return this;
  }

  onError(func: EventType["error"]): this {
    this.listEvents.error = func;
    return this;
  }

  public emitEvent(key: keyof EventType, data?: any) {
    if (this.listEvents[key]) this.listEvents[key](data);
    return this;
  }
}
