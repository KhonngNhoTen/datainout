import * as exceljs from "exceljs";
import { Exporter } from "./Exporter.js";
import { TableData } from "../../common/types/common-type.js";
import { ExceljsExporterHelper } from "../../helpers/exceljs-exporter-helper.js";
import { CellReportOptions, SheetReportOptions } from "../../common/types/report-template.type.js";
import { ExporterOptions } from "../../common/types/exporter.type.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { chunkArray } from "../../helpers/chunk-array.js";
const DEFAULT_CHUNK_SIZE = 100;
export class ExceljsExporter extends Exporter {
  private exporterHelper?: ExceljsExporterHelper;
  private templateManager: ExcelTemplateManager<CellReportOptions> = {} as any;
  private chunksSize: number = 100;
  private useStyle: boolean = true;
  constructor() {
    super(ExceljsExporter.name, "excel");
  }

  async run(data: TableData, options: ExporterOptions & { templateManager: ExcelTemplateManager<CellReportOptions> }): Promise<any> {
    this.templateManager = options.templateManager;
    this.exporterHelper = new ExceljsExporterHelper(options.templateManager);
    this.useStyle = options.useStyle ?? true;

    this.chunksSize = options.chunkSize ?? DEFAULT_CHUNK_SIZE;
    this.templateManager.SheetIndex = 0;

    const workBook = new exceljs.Workbook();
    const workSheet = workBook.addWorksheet();

    const headerTemplate = this.exporterHelper.filterGroupCellDesc("header", this.templateManager.SheetIndex);
    const tableTemplate = this.exporterHelper.filterGroupCellDesc("table", this.templateManager.SheetIndex);
    const footerTemplate = this.exporterHelper.filterGroupCellDesc("footer", this.templateManager.SheetIndex);
    const sheetInformation: SheetReportOptions = this.templateManager.SheetInformation as any;

    // Add header
    if (headerTemplate) this.exporterHelper.setHeader(data.header, headerTemplate, sheetInformation.beginTableAt, workSheet);

    // Add table-content
    console.time("table");
    if (data.table && tableTemplate) {
      if (data.table) data.table = chunkArray(data.table, this.chunksSize);
      data.table.forEach((rows) => {
        this.exporterHelper?.addRows(rows, tableTemplate, workSheet);
      });
    }
    console.timeEnd("table");

    // Add footer
    if (footerTemplate) this.exporterHelper.setFooter(data.footer, footerTemplate, workSheet);

    console.time("merge and width");
    // Merges cells
    if (sheetInformation.merges) this.exporterHelper.mergeCells(workSheet, sheetInformation);

    // Set column width and row height
    if (sheetInformation.columnWidths) this.exporterHelper.setWidthsAndHeights(workSheet, sheetInformation);
    console.timeEnd("merge and width");
    console.time("commit");
    const buffer = await workBook.xlsx.writeBuffer();
    console.timeEnd("commit");
    return buffer;
  }
}
