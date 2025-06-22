import { EventRegister, IBaseStream } from "../../common/core/ListEvents.js";
import { EventType } from "../../common/types/common-type.js";
import { ExporterOptions, ExporterOutputType, ExporterMethodType } from "../../common/types/exporter.type.js";
import { Writable } from "stream";
import { PartialDataTransfer } from "../PartialDataTransfer.js";
import { CellReportOptions } from "../../common/types/report-template.type.js";

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

  addCellTemplate(cells: CellReportOptions[], sheetIndex: number = 0) {
    throw Error("This function only supports for excels");
  }
}

export abstract class ExporterStream extends Exporter implements IBaseStream {
  protected listEvents: EventRegister = new EventRegister();
  protected contents: { header?: any; footer?: any; table: PartialDataTransfer };
  protected templatePath: string;
  protected streamWriter: Writable;
  constructor(
    name: string,
    templatePath: string,
    streamWriter: Writable,
    contents: { header?: any; footer?: any; table: PartialDataTransfer }
  ) {
    super({ name, outputType: "excel" });
    this.contents = contents;
    this.templatePath = templatePath;
    this.streamWriter = streamWriter;
  }

  start(): void {
    const that = this;
    const run = async function () {
      await that.run(that.templatePath, { ...that.contents, stream: that.streamWriter });
      that.contents.table.start((items, hasNext, isNewSheet) => that.add(items as any, hasNext, isNewSheet));
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
