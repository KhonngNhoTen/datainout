import * as exceljs from "exceljs";
import { Exporter } from "./Exporter.js";
import { TableData } from "../../common/types/common-type.js";
import { ExceljsExporterHelper } from "../../helpers/exceljs-exporter-helper.js";

export class ExceljsExporter extends Exporter {
  constructor() {
    super({ name: ExceljsExporter.name, outputType: "excel" });
  }

  async run(templatePath: string, data: TableData): Promise<any> {
    const sheetIndex = 0;
    const exporterHelper = new ExceljsExporterHelper(templatePath);
    const workBook = new exceljs.Workbook();
    const workSheet = workBook.addWorksheet();

    const headerTemplate = exporterHelper.filterGroupCellDesc("header", sheetIndex);
    const tableTemplate = exporterHelper.filterGroupCellDesc("table", sheetIndex);
    const footerTemplate = exporterHelper.filterGroupCellDesc("footer", sheetIndex);
    const sheetInformation = exporterHelper.getSheetInformation(sheetIndex);
    // Add header
    if (headerTemplate) exporterHelper.setHeader(data.header, headerTemplate, sheetInformation.beginTableAt, workSheet);

    // Add table-content
    if (data.table && tableTemplate)
      data.table.forEach((rowdata) => {
        exporterHelper.addRow(rowdata, tableTemplate, workSheet);
      });

    // Add footer
    if (footerTemplate) exporterHelper.setFooter(data.footer, footerTemplate, workSheet);

    // Merges cells
    if (sheetInformation.merges) exporterHelper.mergeCells(workSheet, sheetInformation);

    // Set column width and row height
    if (sheetInformation.columnWidths) exporterHelper.setWidthsAndHeights(workSheet, sheetInformation);
    return await workBook.xlsx.writeBuffer();
  }
}
