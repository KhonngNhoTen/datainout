import * as exceljs from "exceljs";
import { Writable } from "stream";

import { ExporterStream } from "./Exporter.js";
import { ExceljsExporterHelper } from "../../helpers/exceljs-exporter-helper.js";
import { PartialDataTransfer } from "../PartialDataTransfer.js";
import { CellReportOptions, ReportStreamOptions, SheetReportOptions } from "../../common/types/report-template.type.js";
import { ExporterOptions, ExporterStreamOptions } from "../../common/types/exporter.type.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { DEFAULT_END_TABLE } from "../../helpers/excel.helper.js";

export class ExceljsStreamExporter extends ExporterStream {
  private header: any;
  private footer: any;

  private workSheet?: exceljs.Worksheet;
  private workBookWriter?: exceljs.stream.xlsx.WorkbookWriter;

  private exporterHelper?: ExceljsExporterHelper;

  constructor(streamWriter: Writable, options: ExporterStreamOptions) {
    super(ExceljsStreamExporter.name, streamWriter, options);
  }

  async run(
    contents: { header: any; footer: any; stream: Writable },
    options: Omit<ExporterOptions, "templateManager"> & { templateManager: ExcelTemplateManager<CellReportOptions> }
  ): Promise<any> {
    this.templateManager = options.templateManager;
    this.templateManager.SheetIndex = 0;

    this.header = contents.header;
    this.footer = contents.footer;
    this.exporterHelper = new ExceljsExporterHelper(options.templateManager);
    this.workBookWriter = new exceljs.stream.xlsx.WorkbookWriter({
      stream: contents.stream,
      useSharedStrings: this.options?.useSharedStrings,
      useStyles: this.options?.useStyles,
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
      const tableTemplate = this.exporterHelper?.filterGroupCellDesc("table", this.templateManager.SheetIndex);
      if (chunk && tableTemplate) {
        const row = this.exporterHelper?.addRow(chunk, tableTemplate, this.workSheet);
        if (!this.options?.useStyles) row?.commit();
        createdRow++;
      }
    }
    this.listEvents.emitEvent("data", { chunkLenght: chunks.length, createdRow });
    await new Promise((resolve) => setTimeout(resolve, this.options?.sleepTime ?? 10));
  }

  private doneCurrentlySheet() {
    this.listEvents.emitEvent("end", this.workSheet?.name);
    if (this.workSheet) {
      const footerTemplate = this.exporterHelper?.filterGroupCellDesc("footer", this.templateManager.SheetIndex);
      const sheetInformation: SheetReportOptions = this.templateManager.SheetInformation as any;

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
    const sheetName = this.templateManager.SheetInformation.sheetName;
    this.workSheet = this.workBookWriter?.addWorksheet(
      this.templateManager.SheetIndex === 0 ? sheetName : `${sheetName}-${this.templateManager.SheetIndex + 1}`
    );

    this.listEvents.emitEvent("begin", this.workSheet?.name);
    const headerTemplate = this.exporterHelper?.filterGroupCellDesc("header", this.templateManager.SheetIndex);

    if (headerTemplate)
      this.exporterHelper?.setHeader(
        this.header,
        headerTemplate,
        this.templateManager.ActualTableEndRow ?? DEFAULT_END_TABLE,
        this.workSheet as any
      );
    this.templateManager.SheetIndex++;
  }

  private async doneAllSheet() {
    await this.workBookWriter?.commit();
    this.listEvents.emitEvent("finish");
  }
}
