import { ExporterList } from "../type.js";
import { ExcelExporter } from "./Excel.exporter.js";
import { HtmlExporter } from "./Html.exporter.js";
import { PdfExporter } from "./Pdf.exporter.js";

export function exporterFactory(type: ExporterList, templatePath: string, opts?: any) {
  if (type === "excel") return new ExcelExporter(templatePath, opts);
  if (type === "html") return new HtmlExporter(templatePath, opts);
  if (type === "pdf") return new PdfExporter(templatePath, opts);

  throw new Error(`Exporter [${type}] not supports`);
}
