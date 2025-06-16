import { ExporterOptions, ExporterOutputType, ExporterMethodType } from "../../common/types/exporter.type.js";
import { EventType } from "../../common/types/reader.type.js";
import { Writable } from "stream";

export abstract class Exporter {
  protected outputType: ExporterOutputType;
  protected methodType: ExporterMethodType;
  protected name: string;

  protected template: any;

  constructor(opts: ExporterOptions) {
    this.name = opts.name;
    this.outputType = opts.outputType;
    this.methodType = opts?.methodType ?? "full-load";
  }

  abstract run(templatePath: string, data: any): Promise<Buffer | Writable>;
}

export abstract class ExporterStream extends Exporter {
  protected listEvents: Partial<EventType> = {};
  constructor(opts: Omit<ExporterOptions, "methodType">) {
    super({ ...opts, methodType: "stream" });
  }

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

  abstract add(chunks: any[], isNewSheet: boolean): Promise<any>;
}
