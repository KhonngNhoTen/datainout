import ejs from "ejs";
import * as fs from "fs/promises";
import { CreateStreamOpts, ReportData } from "../type.js";
import { Exporter } from "./Exporter.js";
import { WriterStreanm } from "./stream/WriterStream.js";

export class HtmlExporter implements Exporter {
  private templateContent: string = "";
  private templatePath: string = "";

  async setup(templatePath: string): Promise<any> {
    this.templateContent = templatePath;
  }

  async writeFile(reportData: ReportData | ReportData[], path: string) {
    const contentHtml = await this.createContent(reportData);
    await fs.writeFile(path, contentHtml);
  }

  async buffer(reportData: ReportData | ReportData[]): Promise<Buffer> {
    const contentHtml = await this.createContent(reportData);
    return Buffer.from(contentHtml);
  }

  async createContent(reportData: ReportData | ReportData[]) {
    this.templateContent = await fs.readFile(this.templatePath, "utf-8");

    if (!Array.isArray(reportData)) reportData = [reportData];
    const contentHtml = ejs.render(this.templateContent, reportData[0]);
    return contentHtml;
  }

  writerStream(opts: CreateStreamOpts): WriterStreanm {
    throw new Error("Html not support writerStream");
  }
}
