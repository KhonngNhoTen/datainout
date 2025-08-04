import * as ejs from "ejs";
import { IExporter } from "./IExporter.js";

export class HtmlExporter implements IExporter {
  private templatePath: string;
  constructor(templatePath: string) {
    this.templatePath = templatePath;
  }

  async write(reportPath: string, data: any): Promise<void> {
    ejs.render(this.templatePath, data);
  }

  async toBuffer(data: any): Promise<Buffer> {
    return Buffer.from(ejs.render(this.templatePath, data));
  }

  streamTo(...args: any[]): void {
    throw new Error("Stream Method don't support for html.");
  }
}
