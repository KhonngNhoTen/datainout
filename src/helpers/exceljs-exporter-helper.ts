import * as exceljs from "exceljs";
import { SheetSection } from "../common/types/common-type.js";
import { CellReportOptions, SheetReportOptions } from "../common/types/report-template.type.js";
import { getFileExtension } from "./get-file-extension.js";
import { ExcelTemplateManager } from "../common/core/Template.js";

type FilterGroupCellDescOpts = "header" | "table" | "footer";

export class ExceljsExporterHelper {
  templateManager: ExcelTemplateManager<CellReportOptions>;
  useStyle: boolean;
  constructor(templateManager: ExcelTemplateManager<CellReportOptions>, useStyle?: boolean) {
    this.templateManager = templateManager;
    this.useStyle = useStyle ?? false;
  }

  filterGroupCellDesc(opts: FilterGroupCellDescOpts, sheetIndex: number): CellReportOptions[] | null {
    return this.templateManager.GroupCells[opts];
  }

  getSheetInformation() {
    return this.templateManager.SheetInformation;
  }

  setCell(rowData: any, cellOpt: CellReportOptions, cell: exceljs.Cell): exceljs.Cell {
    cell.style = cellOpt.style;
    if (cellOpt.formula) {
      cell.value = {
        formula: cellOpt.formula,
      };
    } else {
      let value = cellOpt.isVariable ? rowData[(cellOpt.value as any).fieldName] : (cellOpt.value as any).hardValue;
      if (cellOpt.formatValue) value = cellOpt.formatValue(value);
      cell.value = value;
    }
    return cell;
  }

  setTitleTable(cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet) {
    const groupTitles = cellsOpts
      .sort((a, b) => b.fullAddress.row - a.fullAddress.row)
      .reduce((acc, value) => {
        const row = value.fullAddress.row;
        acc[row] = acc[row] ? [...acc[row], value] : [value];
        return acc;
      }, {} as Record<number, CellReportOptions[]>);

    const titelTables: exceljs.Row[] = [];
    for (const [rowIndex, footers] of Object.entries(groupTitles)) {
      const row = workSheet.addRow([]);
      footers.forEach((footer) => this.setCell({}, footer, row.getCell(footer.fullAddress.col)));
      titelTables.push(row);
    }
    return titelTables;
  }

  addRow(value: any, cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet): exceljs.Row {
    const row = workSheet.addRow([]);
    for (let i = 0; i < cellsOpts.length; i++) {
      const cellsOpt = cellsOpts[i];
      const cell = row.getCell(cellsOpt.fullAddress.col);
      this.setCell(value, cellsOpt, cell);
    }

    return row;
  }

  addRows(values: any[], cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet): exceljs.Row[] {
    const rows: exceljs.Row[] = values.map((value) => {
      const row = workSheet.addRow([]);
      for (let i = 0; i < cellsOpts.length; i++) {
        const cellsOpt = cellsOpts[i];
        const cell = row.getCell(cellsOpt.fullAddress.col);
        this.setCell(value, cellsOpt, cell);
      }
      if (!this.useStyle) row.commit();
      return row;
    });
    return rows;
  }

  setHeader(header: any, cellsOpts: CellReportOptions[], beginTableAt: number, workSheet: exceljs.Worksheet) {
    const rowHeaders: exceljs.Row[] = [];
    for (let i = 1; i <= beginTableAt; i++) {
      const row = workSheet.addRow([]);
      const formats = cellsOpts.filter((e) => e.fullAddress.row === i);
      formats.forEach((format) => this.setCell(header ?? {}, format, row.getCell(format.fullAddress.col)));
      rowHeaders.push(row);
    }
    return rowHeaders;
  }

  setFooter(value: any, cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet) {
    const groupFooter = cellsOpts
      .sort((a, b) => b.fullAddress.row - a.fullAddress.row)
      .reduce((acc, value) => {
        const row = value.fullAddress.row;
        acc[row] = acc[row] ? [...acc[row], value] : [value];
        return acc;
      }, {} as Record<number, CellReportOptions[]>);

    const _footers: exceljs.Row[] = [];
    for (const [rowIndex, footers] of Object.entries(groupFooter)) {
      const row = workSheet.addRow([]);
      footers.forEach((footer) => this.setCell(value ?? {}, footer, row.getCell(footer.fullAddress.col)));
      _footers.push(row);
    }

    return _footers;
  }

  mergeCells(sheet: exceljs.Worksheet, sheetFormat: Omit<SheetReportOptions, "cells">) {
    if (sheetFormat.merges) {
      const merges = sheetFormat.merges;
      Object.keys(sheetFormat.merges).forEach((masterCell: string) => {
        const { top, left, right, bottom } = merges[masterCell].model;
        sheet.mergeCells(top, left, bottom, right);
      });
    }
  }

  setWidthsAndHeights(sheet: exceljs.Worksheet, sheetFormat: Omit<SheetReportOptions, "cells">) {
    // Set column's width
    sheetFormat.columnWidths?.forEach((colW, i) => {
      if (sheet?.columns && sheet?.columns[i] && colW) sheet.columns[i].width = colW;
    });

    // Set header height
    const rowHeights = sheetFormat.rowHeights;
    Object.keys(rowHeights).forEach((rowIndex, i) => {
      if (rowHeights[rowIndex]) sheet.getRow(rowHeights[i]).height = rowHeights[rowIndex];
    });
  }
}
