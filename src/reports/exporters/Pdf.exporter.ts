import * as fs from "fs/promises";
import { ReportData } from "../type";
import { Exporter } from "./Exporter";
import puppeteer from "puppeteer";
import { HtmlExporter } from "./Html.exporter";

export class PdfExporter implements Exporter {
  private htmlExporter = new HtmlExporter();
  setup(templatePath: string) {
    this.htmlExporter.setup(templatePath);
  }
  async writeFile(reportData: ReportData | ReportData[], path: string): Promise<any> {
    const buffer = await this.createContent(reportData);
    await fs.writeFile(path, buffer);
  }
  async buffer(reportData: ReportData | ReportData[]): Promise<Buffer> {
    return Buffer.from(await this.createContent(reportData));
  }

  private async createContent(reportData: ReportData | ReportData[]): Promise<Uint8Array> {
    const contentHtml = await this.htmlExporter.createContent(reportData);
    const { browser, page } = await this.pagePuppeteerByString(contentHtml);
    const pdf = await page.pdf({
      format: "A4",
    });
    await browser.close();
    return pdf;
  }

  /**
   * Create pdf by puppeteer package.
   * Using a string html converto pdf
   */
  private async pagePuppeteerByString(contentHtml: string) {
    // Start the browser
    const browser = await puppeteer.launch({
      executablePath: "/usr/bin/chromium-browser",
      defaultViewport: {
        width: 1920,
        height: 1080,
      },
    });
    // Open a new blank page
    const page = await browser.newPage();

    // Navigate the page to a URL and wait for everything to load
    // Navigate the page to a URL.
    await page.goto("https://developer.chrome.com/");

    // Use screen CSS instead of print
    await page.emulateMediaType("screen");

    await page.setContent(contentHtml, { waitUntil: "networkidle0" });
    return { page, browser };
  }
}
