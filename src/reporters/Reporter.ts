import * as fs from "fs/promises";
import { Piscina } from "piscina";
import { Writable } from "stream";
import { ExporterOutputType, ExporterStreamOutputType } from "../common/types/exporter.type.js";
import { getConfig } from "../helpers/datainout-config.js";
import { pathReport } from "../helpers/path-file.js";
import { EjsHtmlExporter } from "./exporters/EjsHtml.exporter.js";
import { EjsPdfExporter } from "./exporters/EjsPdf.exporter.js";
import { ExceljsExporter } from "./exporters/Exceljs.exporter.js";
import { Exporter } from "./exporters/Exporter.js";
import { PartialDataTransfer } from "./PartialDataTransfer.js";
import { ExceljsStreamExporter } from "./exporters/ExceljsStream.exporter.js";
import { TableData } from "../common/types/common-type.js";
import { IBaseStream } from "../common/core/ListEvents.js";
import { CellReportOptions, ReportOptions, ReportStreamOptions } from "../common/types/report-template.type.js";
import { ExcelTemplateManager, IExcelTemplateManager } from "../common/core/Template.js";
import { createWorkerPool } from "../common/core/WorkerPool.js";

export class Reporter {
  protected templatePath: string = "";
  protected templateManager: ExcelTemplateManager<CellReportOptions> = {} as any;
  protected type: ExporterOutputType = "excel";
  protected worker?: Piscina<any, any>;

  constructor(templatePath: string, type?: ExporterOutputType) {
    type = type ?? "excel";
    this.templatePath = pathReport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
    if (type === "csv" || type === "excel") {
      this.templateManager = new ExcelTemplateManager(this.templatePath);
      this.templateManager.SheetIndex = 0;
    }
  }

  public get ExcelTemplate(): IExcelTemplateManager<CellReportOptions> {
    if (this.type !== "csv" && this.type !== "excel") throw new Error("ExcelTemplate is only support for excel or csv");
    return this.templateManager;
  }

  async buffer(data: TableData, opts?: ReportOptions): Promise<Buffer>;
  async buffer(data: any, opts?: ReportOptions): Promise<Buffer>;
  async buffer(data: unknown, opts?: ReportOptions): Promise<Buffer> {
    const type = this?.type ?? "excel";
    let exporter: Exporter | undefined;
    if (type === "html") exporter = new EjsHtmlExporter();
    if (type === "excel") exporter = new ExceljsExporter();
    if (type === "pdf") exporter = new EjsPdfExporter();
    if (!exporter) throw new Error("Exporter not setup");

    return (await exporter.run(data, {
      ...(opts ?? {}),
      templatePath: this.templatePath,
      templateManager: this.templateManager,
    })) as Buffer;
  }

  async write(reportPath: string, data: TableData, opts?: ReportOptions): Promise<any>;
  async write(reportPath: string, data: any, opts?: ReportOptions): Promise<any>;
  async write(reportPath: string, data: any, opts?: ReportOptions) {
    const buffer = await this.buffer(data, opts);
    reportPath = pathReport(reportPath, "reportDir");
    await fs.writeFile(data, Uint8Array.from(buffer));
  }

  createStream(
    content: { header?: any; footer?: any; table: PartialDataTransfer },
    streamWriter: Writable,
    opts?: ReportStreamOptions
  ): IBaseStream {
    if (opts?.workerSize) (opts as any).workerPool = createWorkerPool(opts.workerSize);
    const stream = new ExceljsStreamExporter(streamWriter, { ...opts, content, templateManager: this.templateManager });
    return stream;
  }
}
