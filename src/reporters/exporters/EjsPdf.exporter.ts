import * as ejs from "ejs";
import * as puppeteer from "puppeteer";
import { PageData } from "../../common/types/common-type.js";
import { Exporter } from "./Exporter.js";

export class EjsPdfExporter extends Exporter {
  protected override template: string = "";

  constructor() {
    super({ name: EjsPdfExporter.name, outputType: "pdf" });
  }

  async run(templatePath: string, data: PageData): Promise<Buffer> {
    this.template = ejs.render(templatePath, data);
    return this.exportPdf(this.template);
  }

  async exportPdf(html: string) {
    const browser = await puppeteer.launch({
      args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu", "--disable-dev-shm-usage", "--no-zygote"],
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: { top: "20px", bottom: "20px" },
    });
    await browser.close();
    return Buffer.from(pdfBuffer);
  }
}
