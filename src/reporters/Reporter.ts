import { ExporterOutputType } from "../common/types/exporter.type.js";
import { getConfig } from "../helpers/datainout-config.js";
import { pathReport } from "../helpers/path-file.js";
import { CronManager } from "../schedules/CronManager.js";
import { EjsHtmlExporter } from "./exporters/EjsHtml.exporter.js";
import { EjsPdfExporter } from "./exporters/EjsPdf.exporter.js";
import { ExceljsExporter } from "./exporters/Exceljs.exporter.js";
import { Exporter, ExporterStream } from "./exporters/Exporter.js";
import * as fs from "fs/promises";
import { PartialDataTransfer } from "./PartialDataTransfer.js";
import { ExceljsStreamExporter } from "./exporters/ExceljsStream.exporter.js";
import { Writable } from "stream";
export class Reporter {
  protected cronsManager: Array<CronManager> = [];
  protected templatePath: string;

  constructor(templatePath: string) {
    this.templatePath = pathReport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
  }

  async buffer(type: ExporterOutputType, data: any): Promise<Buffer> {
    let exporter: Exporter | undefined;
    if (type === "html") exporter = new EjsHtmlExporter();
    if (type === "excel") exporter = new ExceljsExporter();
    if (type === "pdf") exporter = new EjsPdfExporter();

    if (!exporter) throw new Error("Exporter not setup");
    return (await exporter.run(this.templatePath, data)) as Buffer;
  }

  async write(reportPath: string, type: ExporterOutputType, data: any) {
    const buffer = await this.buffer(type, data);
    reportPath = pathReport(reportPath, "reportDir");
    await fs.writeFile(reportPath, buffer);
  }

  async stream(
    type: ExporterOutputType,
    content: { header?: any; footer?: any; table: PartialDataTransfer },
    streamWriter: Writable
  ): Promise<ExporterStream> {
    const stream = new ExceljsStreamExporter();
    await stream.run(this.templatePath, { stream: streamWriter, footer: content.footer, header: content.header });
    content.table.start(async (items, itemsCount, total, hasNext) => await stream.add(items));
    return stream;
  }
}
