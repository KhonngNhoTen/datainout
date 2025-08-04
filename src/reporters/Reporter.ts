import { getConfig } from "../helpers/datainout-config.js";
import { pathReport } from "../helpers/path-file.js";
import { ExcelExporter } from "./exporters/Excel.exporter.js";
import { HtmlExporter } from "./exporters/Html.exporter.js";
import { PdfExporter } from "./exporters/Pdf.exporter.js";

export class Reporter {
  private templatePath: string;
  private exporter: any;
  constructor(templatePath: string) {
    this.templatePath = pathReport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
  }
  createExporterEXCEL(): ExcelExporter {
    return new ExcelExporter(this.templatePath);
  }
  // createExporterCSV() {}
  createExportePDF() {
    return new PdfExporter(this.templatePath);
  }
  createExporterHTML() {
    return new HtmlExporter(this.templatePath);
  }
}
