import * as exceljs from "exceljs";
import { Exporter } from "./Exporter.js";
import { CellReportOptions, TableReportOptions } from "../../common/types/report-template.type.js";
import {
  setHeader,
  addRow,
  setTitleTable,
  setFooter,
  mergeCells,
  setWidthsAndHeights,
} from "../../helpers/exceljs-report-helper.js";
import { SheetSection, TableData } from "../../common/types/common-type.js";
import { getFileExtension } from "../../helpers/get-file-extension.js";

export class ExceljsExporter extends Exporter {
  protected override template: TableReportOptions = {
    sheets: [],
    name: "",
  };

  private groupCellDescs: { [k in SheetSection]: CellReportOptions[] } = {
    header: [],
    table: [],
    footer: [],
  };

  constructor() {
    super({ name: ExceljsExporter.name, outputType: "excel" });
  }

  async run(templatePath: string, data: TableData): Promise<any> {
    const sheetIndex = 0;
    const workBook = new exceljs.Workbook();
    const workSheet = workBook.addWorksheet();
    if (!workSheet) return;
    this.groupCellDescs = this.getGroupCellDescs(templatePath);

    // add header
    setHeader(
      data?.header,
      this.groupCellDescs.header.filter((e) => e.section === "header"),
      this.template.sheets[0].beginTableAt,
      workSheet
    );

    // add Table
    ///// add title table
    // if (data?.header && this?.groupCellDescs?.header)
    //   setTitleTable(
    //     this.groupCellDescs.header.filter((e) => e.section === "header" && !e.isVariable),
    //     workSheet
    //   );

    ///// add Content
    data?.table?.forEach((raw) => {
      addRow(
        raw,
        this.groupCellDescs.table.filter((e) => e.section === "table" && e.isVariable),
        workSheet
      );
    });

    // add footer
    if (this?.groupCellDescs?.footer && data.footer)
      setFooter(
        data.footer,
        this.groupCellDescs.footer.filter((e) => e.section === "footer"),
        workSheet
      );
    // Add cells in footer-section
    mergeCells(workSheet, this.template.sheets[sheetIndex]);
    setWidthsAndHeights(workSheet, this.template.sheets[sheetIndex]);
    return await workBook.xlsx.writeBuffer();
  }

  getGroupCellDescs(templatePath: string) {
    const sheetIndex = 0;
    this.template = getFileExtension(templatePath) === "js" ? require(templatePath) : require(templatePath).default;
    return this.template.sheets[sheetIndex].cells.reduce((acc, cell) => {
      const section = cell.section ?? "header";
      if (!acc[section]) acc[section] = [cell];
      else acc[section]?.push(cell);
      return acc;
    }, {} as { header: CellReportOptions[]; table: CellReportOptions[]; footer: CellReportOptions[] });
  }
}
