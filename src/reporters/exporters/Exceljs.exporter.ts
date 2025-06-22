import * as exceljs from "exceljs";
import { Exporter } from "./Exporter.js";
import { TableData } from "../../common/types/common-type.js";
import { ExceljsExporterHelper } from "../../helpers/exceljs-exporter-helper.js";
import { CellReportOptions } from "../../common/types/report-template.type.js";

export class ExceljsExporter extends Exporter {
  private exporterHelper?: ExceljsExporterHelper;

  constructor() {
    super({ name: ExceljsExporter.name, outputType: "excel" });
  }

  async run(templatePath: string, data: TableData): Promise<any> {
    const sheetIndex = 0;
    this.exporterHelper = new ExceljsExporterHelper(templatePath);
    const workBook = new exceljs.Workbook();
    const workSheet = workBook.addWorksheet();

    const headerTemplate = this.exporterHelper.filterGroupCellDesc("header", sheetIndex);
    const tableTemplate = this.exporterHelper.filterGroupCellDesc("table", sheetIndex);
    const footerTemplate = this.exporterHelper.filterGroupCellDesc("footer", sheetIndex);
    const sheetInformation = this.exporterHelper.getSheetInformation(sheetIndex);
    // Add header
    if (headerTemplate) this.exporterHelper.setHeader(data.header, headerTemplate, sheetInformation.beginTableAt, workSheet);

    // Add table-content
    if (data.table && tableTemplate)
      data.table.forEach((rowdata) => {
        if (this.exporterHelper) this.exporterHelper.addRow(rowdata, tableTemplate, workSheet);
      });

    // Add footer
    if (footerTemplate) this.exporterHelper.setFooter(data.footer, footerTemplate, workSheet);

    // Merges cells
    if (sheetInformation.merges) this.exporterHelper.mergeCells(workSheet, sheetInformation);

    // Set column width and row height
    if (sheetInformation.columnWidths) this.exporterHelper.setWidthsAndHeights(workSheet, sheetInformation);
    return await workBook.xlsx.writeBuffer();
  }

  addCellTemplate(cells: CellReportOptions[], sheetIndex: number = 0) {
    if (this.exporterHelper) this.addCellTemplate(cells, sheetIndex);
  }
}
