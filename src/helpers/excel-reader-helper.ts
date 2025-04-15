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
};

export class ExcelReaderHelper {
  private beginTable: number = -1;
  private endTableAt: number = -1;
  private rowCount: number = 0;
  private columnIndex = 1;

  private onSheet?: (row: SheetDataHelper) => Promise<any>;
  private onRow?: (row: RowDataHelper) => Promise<any>;
  private onCell?: (cell: CellDataHelper) => Promise<any>;

  constructor(opts?: ExcelReaderHelperOptions) {
    this.onCell = opts?.onCell;
    this.onRow = opts?.onRow;
    this.onSheet = opts?.onSheet;
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
    workBook.eachSheet((sheet, sheetindex) => {
      this.readSheet(sheet, sheetindex);
    });
  }

  private async readSheet(workSheet: exceljs.Worksheet, sheetIndex: number) {
    this.beginTable = -1;
    this.endTableAt = -1;
    this.rowCount = workSheet.rowCount;
    let section;
    for (let i = 1; i <= workSheet.rowCount; i++) {
      const row = workSheet.getRow(i);
      const rowIndex = row.number;
      this.setTable((row.values as any) ?? [], rowIndex);
      for (let j = 0; j < row.cellCount; j++) {
        const cell = row.getCell(j + 1);
        if (cell) {
          const cellData = this.getCell(cell, rowIndex, sheetIndex);
          section = cellData.section;
          if (this.onCell) await this.onCell(cellData);
        }
      }

      if (this.onRow) await this.onRow(this.getRow(row, this.beginTable, this.endTableAt, section));
    }

    if (this.onSheet)
      await this.onSheet({
        beginTableAt: this.beginTable ?? 1,
        sheetIndex,
        name: workSheet.name,
        detail: workSheet,
        endTableAtAt: this.endTableAt,
        rowCount: this.rowCount,
        columnIndex: this.columnIndex,
      });
  }

  public getCell(cell: exceljs.Cell, rowIndex: number, sheetIndex: number): CellDataHelper {
    const section = this.getSection(rowIndex);
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
      beginTableAt: this.beginTable,
      endTableAtAt: this.endTableAt,
      rowCount: this.rowCount,
    };
  }

  public getRow(row: exceljs.Row, beginTableAt: number, endTableAt?: number, section?: SheetSection) {
    return {
      detail: row,
      rowIndex: row.number,
      beginTableAt: beginTableAt,
      endTableAtAt: endTableAt,
      section: section ?? this.getSection(row.number),
    };
  }

  private setTable(values: any[], rowIndex: number) {
    for (let i = 0; i < values.length; i++) {
      const cell = values[i];
      const isVariable = this.isVariable(cell);

      if (isVariable) {
        this.beginTable = this.getBeginTable(cell, rowIndex) ?? this.beginTable;
        this.endTableAt = this.getendTableAt(cell, rowIndex) ?? this.endTableAt;
      } else {
        if (cell && cell.includes(INDEX_COLUMN_TABLE_SYNTAX)) {
          this.columnIndex = i + 1;
          this.beginTable = rowIndex;
        }
      }
    }
  }

  private getBeginTable(cellValue: any, rowIndex: number) {
    cellValue = cellValue + "";
    if (this.isBeginTableNull() && cellValue.includes(VARIABLE_TABLE_SYNTAX)) {
      if (cellValue.includes(INDEX_COLUMN_TABLE_SYNTAX)) return rowIndex;
      return rowIndex;
    }
    return undefined;
  }

  private getendTableAt(cellValue: any, rowIndex: number) {
    cellValue = cellValue + "";
    if (cellValue.includes(END_TABLE_SYNYAX)) return rowIndex - 1;

    if (
      !this.isBeginTableNull() &&
      !cellValue.includes(VARIABLE_TABLE_SYNTAX) &&
      this.isendTableAtNull() &&
      rowIndex - 1 !== this.beginTable
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

  private getSection(rowIndex: number): SheetSection {
    if (this.isBeginTableNull()) return "header";
    if (!this.isBeginTableNull() && this.isendTableAtNull()) return "table";
    if (!this.isendTableAtNull() && rowIndex > this.endTableAt) return "footer";
    return "table";

    // let section: SheetSection = "table";
    // if (rowIndex < this.beginTable) section = "header";
    // else if (!this.isendTableAtNull() && this.rowCount && rowIndex > this.endTableAt) section = "footer";

    // if (rowIndex === 3 || rowIndex === 4) console.log(section);
    // return section;
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

  private isBeginTableNull() {
    return this.beginTable === -1;
  }

  private isendTableAtNull() {
    return this.endTableAt === -1;
  }
}
