import { EventType } from "../../common/types/reader.type.js";
import { BaseReader } from "./BaseReader.js";

export abstract class BaseReaderStream extends BaseReader {
  protected listEvents: Partial<EventType> = {};

  on<EventKey extends keyof EventType>(key: EventKey, func: EventType[EventKey]): this {
    this.listEvents[key] = func;
    return this;
  }

  onFinish(func: EventType["rFinish"]): void {
    this.listEvents.rFinish = func;
  }

  onBegin(func: EventType["rBegin"]): void {
    this.listEvents.rBegin = func;
  }

  onData(func: EventType["rData"]): void {
    this.listEvents.rData = func;
  }

  onEnd(func: EventType["rEnd"]): void {
    this.listEvents.rEnd = func;
  }

  onError(func: EventType["rError"]): void {
    this.listEvents.rError = func;
  }

  protected abstract override load(arg: unknown): Promise<any>;
}
