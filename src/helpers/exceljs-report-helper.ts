import * as exceljs from "exceljs";
import { CellReportOptions, SheetReportOptions } from "../common/types/report-template.type.js";

export function setCell(rowData: any, cellOpt: CellReportOptions, cell: exceljs.Cell): exceljs.Cell {
  cell.style = cellOpt.style;
  cell.value = cellOpt.isVariable ? rowData[(cellOpt.value as any).fieldName] : (cellOpt.value as any).hardValue;
  return cell;
}

export function setTitleTable(cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet) {
  const groupTitles = cellsOpts
    .sort((a, b) => b.fullAddress.row - a.fullAddress.row)
    .reduce((acc, value) => {
      const row = value.fullAddress.row;
      acc[row] = acc[row] ? [...acc[row], value] : [value];
      return acc;
    }, {} as Record<number, CellReportOptions[]>);

  for (const [rowIndex, footers] of Object.entries(groupTitles)) {
    const row = workSheet.addRow([]);
    footers.forEach((footer) => setCell({}, footer, row.getCell(footer.fullAddress.col)));
  }
}

export function addRow(value: any, cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet): exceljs.Worksheet {
  const row = workSheet.addRow([]);
  for (let i = 0; i < cellsOpts.length; i++) {
    const cellsOpt = cellsOpts[i];
    const cell = row.getCell(cellsOpt.fullAddress.col);
    setCell(value, cellsOpt, cell);
  }

  return workSheet;
}

export function setHeader(header: any, cellsOpts: CellReportOptions[], beginTableAt: number, workSheet: exceljs.Worksheet) {
  const rowHeaders = [];
  for (let i = 1; i <= beginTableAt; i++) {
    const row = workSheet.addRow([]);
    const formats = cellsOpts.filter((e) => e.fullAddress.row === i);
    formats.forEach((format) => setCell(header ?? {}, format, row.getCell(format.fullAddress.col)));
    rowHeaders.push(row);
  }
  return workSheet;
}

export function setFooter(value: any, cellsOpts: CellReportOptions[], workSheet: exceljs.Worksheet): exceljs.Worksheet {
  const groupFooter = cellsOpts
    .sort((a, b) => b.fullAddress.row - a.fullAddress.row)
    .reduce((acc, value) => {
      const row = value.fullAddress.row;
      acc[row] = acc[row] ? [...acc[row], value] : [value];
      return acc;
    }, {} as Record<number, CellReportOptions[]>);

  for (const [rowIndex, footers] of Object.entries(groupFooter)) {
    const row = workSheet.addRow([]);
    footers.forEach((footer) => setCell(value ?? {}, footer, row.getCell(footer.fullAddress.col)));
  }

  return workSheet;
}

export function mergeCells(sheet: exceljs.Worksheet, sheetFormat: SheetReportOptions) {
  if (sheetFormat.merges) {
    const merges = sheetFormat.merges;
    Object.keys(sheetFormat.merges).forEach((masterCell: string) => {
      const { top, left, right, bottom } = merges[masterCell].model;
      sheet.mergeCells(top, left, bottom, right);
    });
  }

  return sheet;
}

export function setWidthsAndHeights(sheet: exceljs.Worksheet, sheetFormat: SheetReportOptions) {
  // Set column's width
  sheetFormat.columnWidths?.forEach((colW, i) => {
    if (sheet.columns[i]) sheet.columns[i].width = colW;
  });

  // Set header and footer height
  const rowHeights = sheetFormat.rowHeights;
  Object.keys(rowHeights).forEach((rowIndex) => {
    if (rowHeights[rowIndex]) sheet.getRow(rowHeights[rowIndex]).height = rowHeights[rowIndex];
  });

  return sheet;
}
