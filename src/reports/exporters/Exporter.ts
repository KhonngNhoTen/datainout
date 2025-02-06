import { ReportData, ExporterList } from "../type";
import { ExcelExporter } from "./Excel.exporter";

export interface Exporter {
  setup(templatePath: string): void;
  writeFile(reportData: ReportData | ReportData[], path: string): Promise<any>;
  buffer(reportData: ReportData | ReportData[]): Promise<Buffer>;
}

export function exporterFactory(type: ExporterList) {
  if (type === "excel") return new ExcelExporter();

  throw new Error(`Exporter [${type}] not supports`);
}
