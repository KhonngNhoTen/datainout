import * as exceljs from "exceljs";
import { Writable } from "stream";

import { ExporterStream } from "./Exporter.js";
import { ExceljsExporterHelper } from "../../helpers/exceljs-exporter-helper.js";

export class ExceljsStreamExporter extends ExporterStream {
  private header: any;
  private footer: any;
  private sheetIndex = 0;

  private workSheet?: exceljs.Worksheet;
  private workBookWriter?: exceljs.stream.xlsx.WorkbookWriter;

  private exporterHelper?: ExceljsExporterHelper;

  constructor() {
    super({ name: ExceljsStreamExporter.name, outputType: "excel" });
  }

  async run(templatePath: string, contents: { header: any; footer: any; stream: Writable }): Promise<any> {
    this.header = contents.header;
    this.footer = contents.footer;
    this.exporterHelper = new ExceljsExporterHelper(templatePath);

    this.workBookWriter = new exceljs.stream.xlsx.WorkbookWriter({ stream: contents.stream });
    this.workSheet = this.workBookWriter.addWorksheet();

    if (this.listEvents.rBegin) this.listEvents.rBegin(this.workSheet.name);
    const headerTemplate = this.exporterHelper.filterGroupCellDesc("header", this.sheetIndex);

    if (this.header && headerTemplate)
      this.exporterHelper.setHeader(
        this.header,
        headerTemplate,
        this.exporterHelper.getSheetInformation(this.sheetIndex).beginTableAt,
        this.workSheet
      );
  }

  async add(chunks: any[] | null) {
    if (!this.workSheet) return;

    if (!chunks) {
      this.doneSheet();
      return;
    }

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const tableTemplate = this.exporterHelper?.filterGroupCellDesc("table", this.sheetIndex);
      if (chunk && tableTemplate) this.exporterHelper?.addRow(chunk, tableTemplate, this.workSheet);
      if (this.listEvents.rData) this.listEvents.rData({ section: "table", sheetIndex: 0, sheetName: this.workSheet.name });
    }
  }

  private doneSheet() {
    this.workSheet?.commit();
    if (this.listEvents.rEnd) this.listEvents.rEnd(this.workSheet?.name);
    if (this.workSheet) {
      const footerTemplate = this.exporterHelper?.filterGroupCellDesc("footer", this.sheetIndex);
      const sheetInformation = this.exporterHelper?.getSheetInformation(this.sheetIndex);

      // Set footer
      if (this.footer && footerTemplate && this.exporterHelper) this.exporterHelper.setFooter(this.footer, footerTemplate, this.workSheet);

      // Merges cells
      if (this.exporterHelper && sheetInformation?.merges) this.exporterHelper.mergeCells(this.workSheet, sheetInformation);

      // Set heigth and width
      if (this.exporterHelper && sheetInformation?.columnWidths)
        this.exporterHelper.setWidthsAndHeights(this.workSheet, this.template.sheets[0]);
    }
    this.doneAllSheet();
  }

  private doneAllSheet() {
    this.workBookWriter?.commit();
    if (this.listEvents.rFinish) this.listEvents.rFinish();
  }
}
