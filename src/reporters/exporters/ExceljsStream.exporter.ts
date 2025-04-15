import * as exceljs from "exceljs";
import { PassThrough, Writable } from "stream";

import { SheetSection } from "../../common/types/common-type.js";
import { CellReportOptions, TableReportOptions } from "../../common/types/report-template.type.js";
import { getFileExtension } from "../../helpers/get-file-extension.js";
import { ExporterStream } from "./Exporter.js";
import {
  addRow,
  mergeCells,
  setFooter,
  setHeader,
  setTitleTable,
  setWidthsAndHeights,
} from "../../helpers/exceljs-report-helper.js";

export class ExceljsStreamExporter extends ExporterStream {
  protected override template: TableReportOptions = {
    sheets: [],
    name: "",
  };

  private groupCellDescs: { [k in SheetSection]: CellReportOptions[] } = {
    header: [],
    table: [],
    footer: [],
  };

  private contents: { header: any; footer: any } = {
    header: undefined,
    footer: undefined,
  };

  private workSheet?: exceljs.Worksheet;
  private workBookWriter?: exceljs.stream.xlsx.WorkbookWriter;

  constructor() {
    super({ name: ExceljsStreamExporter.name, outputType: "excel" });
  }

  async run(templatePath: string, contents: { header: any; footer: any; stream: Writable }): Promise<any> {
    const sheetIndex = 0;

    this.contents = { footer: contents.footer, header: contents.header };
    this.workBookWriter = new exceljs.stream.xlsx.WorkbookWriter({ stream: contents.stream });
    this.workSheet = this.workBookWriter.addWorksheet();
    this.groupCellDescs = this.getGroupCellDescs(templatePath);

    if (this.listEvents.rBegin) this.listEvents.rBegin(this.workSheet.name);

    setHeader(contents.header, this.groupCellDescs.header, this.template.sheets[sheetIndex].beginTableAt, this.workSheet);
    setTitleTable(
      this.groupCellDescs.table.filter((e) => !e.isVariable),
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
      addRow(
        chunk,
        this.groupCellDescs.table.filter((e) => e.isVariable),
        this.workSheet
      );

      if (this.listEvents.rData) this.listEvents.rData({ section: "table", sheetIndex: 0, sheetName: this.workSheet.name });
    }
  }

  private getGroupCellDescs(templatePath: string) {
    const sheetIndex = 0;
    this.template = getFileExtension(templatePath) === "js" ? require(templatePath) : require(templatePath).default;
    return this.template.sheets[sheetIndex].cells.reduce((acc, cell) => {
      const section = cell.section ?? "header";
      if (!acc[section]) acc[section] = [cell];
      else acc[section]?.push(cell);
      return acc;
    }, {} as { header: CellReportOptions[]; table: CellReportOptions[]; footer: CellReportOptions[] });
  }

  private doneSheet() {
    this.workSheet?.commit();
    if (this.listEvents.rEnd) this.listEvents.rEnd(this.workSheet?.name);
    if (this.workSheet) {
      setFooter(this.contents.footer, this.groupCellDescs.footer, this.workSheet);
      mergeCells(this.workSheet, this.template.sheets[0]);
      setWidthsAndHeights(this.workSheet, this.template.sheets[0]);
    }
    this.doneAllSheet();
  }

  private doneAllSheet() {
    this.workBookWriter?.commit();
    if (this.listEvents.rFinish) this.listEvents.rFinish();
  }
}
