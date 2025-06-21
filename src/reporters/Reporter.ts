import * as fs from "fs/promises";
import { Writable } from "stream";
import { ExporterOutputType, ExporterStreamOutputType } from "../common/types/exporter.type.js";
import { getConfig } from "../helpers/datainout-config.js";
import { pathReport } from "../helpers/path-file.js";
import { EjsHtmlExporter } from "./exporters/EjsHtml.exporter.js";
import { EjsPdfExporter } from "./exporters/EjsPdf.exporter.js";
import { ExceljsExporter } from "./exporters/Exceljs.exporter.js";
import { Exporter, ExporterStream } from "./exporters/Exporter.js";
import { PartialDataTransfer } from "./PartialDataTransfer.js";
import { ExceljsStreamExporter } from "./exporters/ExceljsStream.exporter.js";
import { TableData } from "../common/types/common-type.js";
import { IBaseStream } from "../common/core/ListEvents.js";
export class Reporter {
  protected templatePath: string;

  constructor(templatePath: string) {
    this.templatePath = pathReport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
  }

  async buffer(data: TableData, type?: ExporterOutputType): Promise<Buffer>;
  async buffer(data: any, type?: ExporterOutputType): Promise<Buffer>;
  async buffer(data: unknown, type?: ExporterOutputType): Promise<Buffer> {
    type = type ?? "excel";
    let exporter: Exporter | undefined;
    if (type === "html") exporter = new EjsHtmlExporter();
    if (type === "excel") exporter = new ExceljsExporter();
    if (type === "pdf") exporter = new EjsPdfExporter();
    if (!exporter) throw new Error("Exporter not setup");
    return (await exporter.run(this.templatePath, data)) as Buffer;
  }

  async write(reportPath: string, data: TableData, type?: ExporterOutputType): Promise<any>;
  async write(reportPath: string, data: any, type?: ExporterOutputType): Promise<any>;
  async write(reportPath: string, data: any, type?: ExporterOutputType) {
    type = type ?? "excel";
    const buffer = await this.buffer(data, type);
    reportPath = pathReport(reportPath, "reportDir");
    await fs.writeFile(reportPath, Uint8Array.from(buffer));
  }

  createStream(
    content: { header?: any; footer?: any; table: PartialDataTransfer },
    streamWriter: Writable,
    type: ExporterStreamOutputType = "excel"
  ): IBaseStream {
    const stream = new ExceljsStreamExporter(this.templatePath, streamWriter, {
      footer: content.footer,
      header: content.header,
      table: content.table,
    });
    return stream;
  }
}
