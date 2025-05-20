import * as exceljs from "exceljs";
// import { CellImportOptions, SheetImportOptions } from "../common/types/import-template.type.js";
import { AttributeType, SheetSection } from "../common/types/common-type.js";
// const TITLE_TABLE_SYNTAX = "$$[title]";
const END_TABLE_SYNYAX = "$$[br]";
const VARIABLE_TABLE_SYNTAX = "$$";
const VARIABLE_SYNTAX = "$";
const INDEX_COLUMN_TABLE_SYNTAX = "$$*";

export type CellDataHelper = {
  sheetIndex: number;
  rowIndex: number;
  isVariable: boolean;
  label?: string;
  variableValue?: { fieldName: string; type: AttributeType };
  detail: exceljs.Cell;
  address: string;
  section: SheetSection;
  beginTableAt?: number;
  endTableAtAt?: number;
  rowCount: number;
};

export type RowDataHelper = {
  rowIndex: number;
  detail: exceljs.Row;
  beginTableAt?: number;
  endTableAtAt?: number;
  section: SheetSection;
};

export type SheetDataHelper = {
  sheetIndex: number;
  columnIndex: number;
  name: string;
  detail: exceljs.Worksheet;
  beginTableAt: number;
  endTableAtAt?: number;
  rowCount: number;
};

type ExcelReaderHelperOptions = {
  onSheet?: (sheet: SheetDataHelper) => Promise<any>;
  onRow?: (row: RowDataHelper) => Promise<any>;
  onCell?: (cell: CellDataHelper) => Promise<any>;
  beginTables?: number[];
  endTableAts?: (number | undefined)[];
  columnIndexes?: number[];
};

export class ExcelReaderHelper {
  // private beginTable: number = -1;
  // private endTableAt: number = -1;
  // private columnIndex = 1;
  private rowCount: number = 0;

  private beginTables: number[] = [];
  private endTableAts: (number | undefined)[] = [];
  private columnIndexes: number[] = [];

  private onSheet?: (row: SheetDataHelper) => Promise<any>;
  private onRow?: (row: RowDataHelper) => Promise<any>;
  private onCell?: (cell: CellDataHelper) => Promise<any>;

  constructor(opts?: ExcelReaderHelperOptions) {
    this.onCell = opts?.onCell;
    this.onRow = opts?.onRow;
    this.onSheet = opts?.onSheet;
    this.beginTables = opts?.beginTables ?? [];
    this.endTableAts = opts?.endTableAts ?? [];
    this.columnIndexes = opts?.columnIndexes ?? [];
  }

  async load(file: string): Promise<any>;
  async load(buffer: Buffer): Promise<any>;
  async load(arg: unknown) {
    await this.readWorkBook(arg as string | Buffer);
  }

  private async readWorkBook(file: string | Buffer) {
    const workBook = new exceljs.Workbook();
    if (file instanceof Buffer) await workBook.xlsx.load(file as unknown as exceljs.Buffer);
    else await workBook.xlsx.readFile(file);
    // workBook.eachSheet((sheet, sheetindex) => {
    //   this.readSheet(sheet, sheetindex);
    // });

    for (let i = 0; i < workBook.worksheets.length; i++) {
      const workSheet = workBook.getWorksheet(i + 1);
      if (workSheet) await this.readSheet(workSheet, i + 1);
    }
  }

  private async readSheet(workSheet: exceljs.Worksheet, sheetIndex: number) {
    let beginTable = this.beginTables[sheetIndex] ?? -1;
    let endTableAt = this.endTableAts[sheetIndex] ?? -1;
    let columnIndex = this.columnIndexes[sheetIndex] ?? -1;

    this.rowCount = workSheet.rowCount;
    let section;
    for (let i = 1; i <= workSheet.rowCount; i++) {
      const row = workSheet.getRow(i);
      const rowIndex = row.number;
      const tables = this.getTableInfor((row.values as any) ?? [], rowIndex, beginTable, endTableAt);
      beginTable = tables.beginTable ?? beginTable;
      endTableAt = tables.endTableAt ?? endTableAt;
      columnIndex = tables.columnIndex ?? columnIndex;

      section = this.getSection(row.number, sheetIndex, row.values as any[], beginTable, endTableAt, columnIndex);

      for (let j = 0; j < row.cellCount; j++) {
        const cell = row.getCell(j + 1);
        if (cell) {
          const cellData = this.getCell(cell, rowIndex, sheetIndex, section, beginTable, endTableAt, columnIndex);
          section = cellData.section;
          if (this.onCell) await this.onCell(cellData);
        }
      }

      if (this.onRow) await this.onRow(this.getRow(row, sheetIndex, section, beginTable, endTableAt, columnIndex));
    }

    if (this.onSheet)
      await this.onSheet({
        beginTableAt: beginTable ?? 1,
        sheetIndex,
        name: workSheet.name,
        detail: workSheet,
        endTableAtAt: endTableAt,
        rowCount: this.rowCount,
        columnIndex: columnIndex,
      });
  }

  public getCell(
    cell: exceljs.Cell,
    rowIndex: number,
    sheetIndex: number,
    section: SheetSection,
    beginTableAt?: number,
    endTableAt?: number,
    columnIndex?: number
  ): CellDataHelper {
    let _beginTableAt = beginTableAt ?? this.beginTables[sheetIndex] ?? -1,
      _endTableAt = endTableAt ?? this.endTableAts[sheetIndex] ?? -1,
      _columnIndex = columnIndex ?? this.columnIndexes[sheetIndex] ?? -1;

    const isVariable = this.isVariable(cell.value);
    return {
      address: this.getAddress(cell.address, section, isVariable),
      detail: cell,
      isVariable,
      rowIndex,
      section,
      sheetIndex,
      label: this.getLabel(cell.value + "", isVariable),
      variableValue: isVariable ? this.getVariableValue(cell) : undefined,
      beginTableAt: _beginTableAt,
      endTableAtAt: _endTableAt,
      rowCount: this.rowCount,
    };
  }

  public getRow(
    row: exceljs.Row,
    sheetIndex: number,
    section: SheetSection,
    beginTableAt?: number,
    endTableAt?: number,
    columnIndex?: number
  ) {
    let _beginTableAt = beginTableAt ?? this.beginTables[sheetIndex] ?? -1,
      _endTableAt = endTableAt ?? this.endTableAts[sheetIndex] ?? -1,
      _columnIndex = columnIndex ?? this.columnIndexes[sheetIndex] ?? -1;
    return {
      detail: row,
      rowIndex: row.number,
      beginTableAt: _beginTableAt,
      endTableAtAt: _endTableAt,
      section: section,
    };
  }

  getSection(
    rowIndex: number,
    sheetIndex: number,
    cellValues: any[],
    beginTableAt?: number,
    endTableAt?: number,
    columnIndex?: number
  ): SheetSection {
    let _beginTableAt = beginTableAt ?? this.beginTables[sheetIndex] ?? -1,
      _endTableAt = endTableAt ?? this.endTableAts[sheetIndex] ?? -1,
      _columnIndex = columnIndex ?? this.columnIndexes[sheetIndex] ?? -1;

    if (this.isBeginTableNull(_beginTableAt)) return "header";
    if (rowIndex < _beginTableAt) return "header";
    if (rowIndex > _beginTableAt && cellValues[_columnIndex - 1]) return "table";
    if (rowIndex > _endTableAt && !cellValues[_columnIndex - 1]) return "footer";
    return "header";
  }

  private getTableInfor(values: any[], rowIndex: number, beginTable: number, endTableAt: number) {
    let columnIndex;
    for (let i = 0; i < values.length; i++) {
      const cell = values[i];
      const isVariable = this.isVariable(cell);

      if (isVariable) {
        beginTable = this.getBeginTable(cell, rowIndex, beginTable) ?? beginTable;
        endTableAt = this.getEndTableAt(cell, rowIndex, beginTable, endTableAt) ?? endTableAt;
      } else {
        if (cell && (cell + "").includes(INDEX_COLUMN_TABLE_SYNTAX)) {
          columnIndex = i + 1;
          beginTable = rowIndex;
        }
      }
    }

    return { beginTable, endTableAt, columnIndex };
  }

  private getBeginTable(cellValue: any, rowIndex: number, beginTable: number) {
    cellValue = cellValue + "";
    if (this.isBeginTableNull(beginTable) && cellValue.includes(VARIABLE_TABLE_SYNTAX)) {
      if (cellValue.includes(INDEX_COLUMN_TABLE_SYNTAX)) return rowIndex;
      return rowIndex;
    }
    return undefined;
  }

  private getEndTableAt(cellValue: any, rowIndex: number, beginTable: number, endTableAt: number) {
    cellValue = cellValue + "";
    if (cellValue.includes(END_TABLE_SYNYAX)) return rowIndex - 1;

    if (
      !this.isBeginTableNull(beginTable) &&
      !cellValue.includes(VARIABLE_TABLE_SYNTAX) &&
      this.isEndTableAtNull(endTableAt) &&
      rowIndex - 1 !== beginTable
    )
      return rowIndex - 1;
    // return rowIndex - 1 - this.rowCount;
    return undefined;
  }

  private isVariable(cellValue: any) {
    cellValue = cellValue + "";
    if (cellValue.includes(INDEX_COLUMN_TABLE_SYNTAX)) return false;
    return cellValue.includes(VARIABLE_SYNTAX);
  }

  private getLabel(label: string, isVariable: boolean) {
    if (isVariable) return undefined;
    if (label.includes(INDEX_COLUMN_TABLE_SYNTAX)) return label.split("->").pop();
    return label;
  }

  private getAddress(address: string, section: SheetSection, isVariable: boolean) {
    return !isVariable ? address : section !== "header" ? this.getColumnIndex(address) : address;
  }

  private getColumnIndex(address: string) {
    return address.split(/\d+/)[0];
  }

  private getVariableValue(cell: exceljs.Cell): CellDataHelper["variableValue"] {
    const cellValue = (cell.value + "").replace("$$", "$");

    let fieldName = "";
    let type: AttributeType = "string";
    fieldName = cellValue.split("$")[1];
    if (fieldName.includes("->")) {
      const args = fieldName.split("->");
      fieldName = args[0];
      type = args[1].toLowerCase() as AttributeType;
    }

    return { fieldName, type };
  }

  private isBeginTableNull(beginTable: number) {
    return beginTable === -1;
  }

  private isEndTableAtNull(endTableAt: number) {
    return endTableAt === -1;
  }
}
