import * as ejs from "ejs";
import * as puppeteer from "puppeteer";
import { PageData } from "../../common/types/common-type.js";
import { Exporter } from "./Exporter.js";
import { ExporterOptions } from "../../common/types/exporter.type.js";

export class EjsPdfExporter extends Exporter {
  protected template: string = "";

  constructor() {
    super(EjsPdfExporter.name, "pdf");
  }

  async run(data: PageData, options: ExporterOptions): Promise<Buffer> {
    this.template = ejs.render(options?.templatePath ?? "", data);
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
