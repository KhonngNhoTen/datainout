import * as ejs from "ejs";
import * as fs from "fs";
import * as puppeteer from "puppeteer";
import { IExporter } from "./IExporter.js";

export class PdfExporter implements IExporter {
  private templatePath: string;
  constructor(templatePath: string) {
    this.templatePath = templatePath;
  }

  async write(reportPath: string, data: any, opts?: any): Promise<void> {
    const { browser, content } = await this.genTemplate(data, opts);
    await browser.close();
    fs.writeFile(reportPath, content, () => console.log("Write file pdf successfully"));
  }

  async toBuffer(data: any, opts?: any): Promise<Buffer> {
    const { browser, content } = await this.genTemplate(data, opts);
    await browser.close();
    return content as unknown as Buffer;
  }

  private async genTemplate(data: any, opts?: any) {
    const html = ejs.render(this.templatePath ?? "", data);
    const browser = await puppeteer.launch({
      args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu", "--disable-dev-shm-usage", "--no-zygote"],
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    return {
      content: await page.pdf({
        format: "A4",
        printBackground: true,
        margin: { top: "20px", bottom: "20px" },
      }),
      browser,
    };
  }
  streamTo(...args: any[]): void {
    throw new Error("Method not implemented.");
  }
}
