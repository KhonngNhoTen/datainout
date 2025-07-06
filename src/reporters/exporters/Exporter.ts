import { EventRegister, IBaseStream } from "../../common/core/ListEvents.js";
import { EventType } from "../../common/types/common-type.js";
import { ExporterOutputType, ExporterMethodType, ExporterOptions, ExporterStreamOptions } from "../../common/types/exporter.type.js";
import { Writable } from "stream";
import { PartialDataTransfer } from "../PartialDataTransfer.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { CellReportOptions, ReportStreamOptions } from "../../common/types/report-template.type.js";

export abstract class Exporter {
  protected outputType: ExporterOutputType;
  protected methodType: ExporterMethodType;
  protected name: string;
  protected options?: any;

  constructor(name: string, outputType: ExporterOutputType, methodType?: ExporterMethodType) {
    this.name = name;
    this.outputType = outputType;
    this.methodType = methodType ?? "full-load";
  }

  abstract run(data: any, options?: ExporterOptions): Promise<Buffer | Writable>;
}

export abstract class ExporterStream extends Exporter implements IBaseStream {
  protected listEvents: EventRegister = new EventRegister();
  protected templateManager: ExcelTemplateManager<CellReportOptions> = {} as any;
  protected streamWriter: Writable;
  protected override options: ExporterStreamOptions = {} as any;
  constructor(name: string, streamWriter: Writable, options: ExporterStreamOptions) {
    super(name, "excel");
    this.streamWriter = streamWriter;
    this.options = options;
  }

  start(): void {
    const that = this;
    const run = async function () {
      await that.run(that.options.content, that.options);
      that.options.content.table.start(async (items, hasNext, isNewSheet) => await that.add(items as any, hasNext, isNewSheet));
    };
    run();
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
  abstract add(chunks: any[], hasNext: boolean, isNewSheet: boolean): Promise<any>;
}
