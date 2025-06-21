import * as exceljs from "exceljs";
import { SheetSection } from "../common/types/common-type.js";
import { CellReportOptions, SheetReportOptions, TableReportOptions } from "../common/types/report-template.type.js";
import { getFileExtension } from "./get-file-extension.js";

type FilterGroupCellDescOpts = "header" | "table" | "footer";

export class ExceljsExporterHelper {
  groupCellDescs: { [k in SheetSection]: CellReportOptions[] }[] = [];
  sheetsInfor: Omit<SheetReportOptions, "cells">[] = [];

  constructor(templatePath: string) {
    const { groupCells, sheetInformation } = this.getGroupCellDescs(templatePath);
    this.groupCellDescs = groupCells;
    this.sheetsInfor = sheetInformation;
  }

  filterGroupCellDesc(opts: FilterGroupCellDescOpts, sheetIndex: number): CellReportOptions[] | null {
    const sheetTemplate = this.groupCellDescs[sheetIndex];
    return sheetTemplate[opts];
  }

  getSheetInformation(sheetIndex: number) {
    return this.sheetsInfor[sheetIndex];
  }

  getGroupCellDescs(templatePath: string) {
    const template: TableReportOptions = getFileExtension(templatePath) === "js" ? require(templatePath) : require(templatePath).default;
    const groupCells: { [k in SheetSection]: CellReportOptions[] }[] = [];
    const sheetInformation: Omit<SheetReportOptions, "cells">[] = [];

    template.sheets.forEach((sheet) => {
      sheetInformation.push({
        beginTableAt: sheet.beginTableAt,
        endTableAt: sheet.endTableAt,
        keyTableAt: sheet.keyTableAt,
        sheetIndex: sheet.sheetIndex,
        sheetName: sheet.sheetName,
        rowHeights: sheet.rowHeights,
        columnWidths: sheet.columnWidths,
        merges: sheet.merges,
        pageSize: sheet.pageSize,
      });
      const cellsDes: any = {};
      sheet.cells.forEach((cell) => {
        const section = cell.section ?? "header";
        if (!cellsDes[section]) cellsDes[section] = [cell];
        else cellsDes[section]?.push(cell);
      });
      groupCells.push(cellsDes);
    });

    return { sheetInformation, groupCells };
  }

  setCell(rowData: any, cellOpt: CellReportOptions, cell: exceljs.Cell): exceljs.Cell {
    cell.style = cellOpt.style;
    if (cellOpt.formula) {
      cell.value = {
        formula: cellOpt.formula,
      };
    } else cell.value = cellOpt.isVariable ? rowData[(cellOpt.value as any).fieldName] : (cellOpt.value as any).hardValue;
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
