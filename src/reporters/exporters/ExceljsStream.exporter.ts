import * as exceljs from "exceljs";
import { Writable } from "stream";

import { ExporterStream } from "./Exporter.js";
import { ExceljsExporterHelper } from "../../helpers/exceljs-exporter-helper.js";
import { PartialDataTransfer } from "../PartialDataTransfer.js";
import { CellReportOptions, ReportStreamOptions } from "../../common/types/report-template.type.js";

export class ExceljsStreamExporter extends ExporterStream {
  private header: any;
  private footer: any;
  private sheetIndex = 0;
  private sheetDescIndex = 0;

  private workSheet?: exceljs.Worksheet;
  private workBookWriter?: exceljs.stream.xlsx.WorkbookWriter;

  private exporterHelper?: ExceljsExporterHelper;
  private opts?: ReportStreamOptions;

  constructor(
    templatePath: string,
    streamWriter: Writable,
    contents: { header?: any; footer?: any; table: PartialDataTransfer },
    opts?: ReportStreamOptions
  ) {
    super(ExceljsStreamExporter.name, templatePath, streamWriter, contents);
    this.opts = opts;
  }

  async run(templatePath: string, contents: { header: any; footer: any; stream: Writable }): Promise<any> {
    this.header = contents.header;
    this.footer = contents.footer;
    this.exporterHelper = new ExceljsExporterHelper(templatePath);
    this.workBookWriter = new exceljs.stream.xlsx.WorkbookWriter({
      stream: contents.stream,
      useSharedStrings: this.opts?.useSharedStrings,
      useStyles: this.opts?.useStyles,
    });

    this.createSheet();
  }

  async add(chunks: any[] | null, hasNext: boolean, isNewSheet: boolean = false) {
    if (!this.workSheet) return;
    if (!chunks || hasNext === false) {
      this.doneCurrentlySheet();
      await this.doneAllSheet();
      return;
    }

    if (isNewSheet) {
      this.doneCurrentlySheet();
      this.createSheet();
    }

    let createdRow = 0;
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const tableTemplate = this.exporterHelper?.filterGroupCellDesc("table", this.sheetDescIndex);
      if (chunk && tableTemplate) {
        const row = this.exporterHelper?.addRow(chunk, tableTemplate, this.workSheet);
        if (!this.opts?.useStyles) row?.commit();
        createdRow++;
      }
    }
    this.listEvents.emitEvent("data", { chunkLenght: chunks.length, createdRow });
    await new Promise((resolve) => setTimeout(resolve, this.opts?.sleepTime ?? 10));
  }

  private doneCurrentlySheet() {
    this.listEvents.emitEvent("end", this.workSheet?.name);
    if (this.workSheet) {
      const footerTemplate = this.exporterHelper?.filterGroupCellDesc("footer", this.sheetDescIndex);
      const sheetInformation = this.exporterHelper?.getSheetInformation(this.sheetDescIndex);

      // Set footer
      if (footerTemplate && this.exporterHelper) this.exporterHelper.setFooter(this.footer ?? {}, footerTemplate, this.workSheet);

      // Merges cells
      if (this.exporterHelper && sheetInformation?.merges) this.exporterHelper.mergeCells(this.workSheet, sheetInformation);

      // Set heigth and width
      if (this.exporterHelper && sheetInformation?.columnWidths) this.exporterHelper.setWidthsAndHeights(this.workSheet, sheetInformation);
    }
    this.workSheet?.commit();
  }

  private createSheet() {
    const sheetName = this.exporterHelper?.getSheetInformation(this.sheetIndex).sheetName;
    this.workSheet = this.workBookWriter?.addWorksheet(this.sheetIndex === 0 ? sheetName : `${sheetName}-${this.sheetIndex + 1}`);

    this.listEvents.emitEvent("begin", this.workSheet?.name);
    const headerTemplate = this.exporterHelper?.filterGroupCellDesc("header", this.sheetDescIndex);

    if (headerTemplate)
      this.exporterHelper?.setHeader(
        this.header,
        headerTemplate,
        this.exporterHelper.getSheetInformation(this.sheetIndex).beginTableAt,
        this.workSheet as any
      );
    this.sheetIndex++;
  }

  private async doneAllSheet() {
    await this.workBookWriter?.commit();
    this.listEvents.emitEvent("finish");
  }

  addCellTemplate(cells: CellReportOptions[], sheetIndex: number = 0) {
    if (this.exporterHelper) this.addCellTemplate(cells, sheetIndex);
  }
}
