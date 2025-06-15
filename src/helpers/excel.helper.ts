import { AttributeType, SheetExcelOption, SheetSection } from "../common/types/common-type.js";
import * as exceljs from "exceljs";
import { CellDataHelper, ExcelReaderHelperOptions, RowDataHelper, SheetDataHelper } from "../common/types/excel-reader-helper.type.js";

export const SYNTAX = {
  VARIABLE_TABLE_SYNTAX: "$$",
  VARIABLE_SYNTAX: "$",
  INDEX_COLUMN_TABLE_SYNTAX: "$$**",
  END_TABLE_SYNYAX: "$$br",
};

export const DEFAULT_BEGIN_TABLE = -1;
export const DEFAULT_END_TABLE = -1;
export const DEFAULT_COLUMN_INDEX = 1;

export class ReaderExceljsHelper {
  private isSampleExcel: boolean = true;
  private template: SheetExcelOption[] = [];

  private onSheet?: (row: SheetDataHelper) => Promise<any>;
  private onRow?: (row: RowDataHelper) => Promise<any>;
  private onCell?: (cell: CellDataHelper) => Promise<any>;

  constructor(opts?: ExcelReaderHelperOptions) {
    this.onCell = opts?.onCell;
    this.onRow = opts?.onRow;
    this.onSheet = opts?.onSheet;
    this.isSampleExcel = opts?.isSampleExcel ?? true;
    this.template = opts?.template ?? [];
  }

  async load(file: string): Promise<any>;
  async load(buffer: Buffer): Promise<any>;
  async load(arg: unknown) {
    const workBook = new exceljs.Workbook();
    if (arg instanceof Buffer) await workBook.xlsx.load(arg as unknown as exceljs.Buffer);
    else await workBook.xlsx.readFile(arg as string);

    for (let i = 0; i < workBook.worksheets.length; i++) {
      const workSheet = workBook.getWorksheet(i + 1);
      if (workSheet) await this.eachSheet(workSheet, i + 1);
    }
  }

  private async eachSheet(sheet: exceljs.Worksheet, sheetIndex: number) {
    const sheetDesc = this.template[sheetIndex - 1] ?? [];
    let beginTable = DEFAULT_BEGIN_TABLE;
    let endTable = DEFAULT_END_TABLE;
    let columnIndex = DEFAULT_COLUMN_INDEX;

    for (let i = 1; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      beginTable = ReaderExceljsHelper.beginTableAt(row, sheetDesc, this.isSampleExcel) ?? beginTable;
      endTable = ReaderExceljsHelper.endTableAt(row, sheetDesc, this.isSampleExcel) ?? endTable;
      columnIndex = ReaderExceljsHelper.columnTableIndex(row, sheetDesc, this.isSampleExcel) ?? columnIndex;

      const section = ReaderExceljsHelper.getSection(row, beginTable, endTable);
      const cells: CellDataHelper[] = [];
      if (this.onCell)
        for (let j = 0; j < row.cellCount; j++) {
          const cell = this.convertCell(row.getCell(j + 1), section, beginTable, endTable);
          cells.push(cell);
          this.onCell(cell);
        }

      if (this.onRow) await this.onRow(this.convertRow(row, section, cells));
    }

    if (this.onSheet)
      await this.onSheet({
        beginTableAt: beginTable ?? 1,
        columnIndex,
        endTableAt: endTable,
        sheetIndex,
        name: sheet.name,
        detail: sheet,
        rowCount: sheet.rowCount,
      });
  }

  private convertCell(cell: exceljs.Cell, section: SheetSection, beginTableAt: number, endTableAt: number) {
    const isVariable = ReaderExceljsHelper.isVariable(cell.value);
    return {
      address: ReaderExceljsHelper.getAddress(cell.address, section, isVariable),
      detail: cell,
      isVariable,
      rowIndex: +cell.fullAddress.row,
      section,
      label: ReaderExceljsHelper.getLabel(cell.value + "", isVariable),
      variableValue: isVariable ? ReaderExceljsHelper.getVariableValue(cell) : undefined,
      beginTableAt,
      endTableAt,
      formula: cell.formula ?? undefined,
    };
  }

  private convertRow(row: exceljs.Row, section: SheetSection, cells: CellDataHelper[]) {
    return {
      detail: row,
      rowIndex: row.number,
      section: section,
      cells,
      beginTableAt: cells[0].beginTableAt ?? DEFAULT_BEGIN_TABLE,
      endTableAt: cells[0].endTableAt ?? DEFAULT_END_TABLE,
    };
  }

  static getSection(row: exceljs.Row, beginTable: number, endTable: number): SheetSection {
    if (beginTable === DEFAULT_BEGIN_TABLE) return "header";
    if (row.number > endTable && endTable !== DEFAULT_END_TABLE) return "footer";
    if (row.number > beginTable && row.number) return "table";
    return "header";
  }

  /** Get begin table at by row */
  static beginTableAt(row: exceljs.Row, sheetOpts: SheetExcelOption, isSampleExcel: boolean = true) {
    if (isSampleExcel)
      for (let i = 0; i < row.cellCount; i++) {
        const cellValue = row.getCell(i + 1) as unknown as string;
        if (typeof cellValue === "string" && cellValue.includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX)) return row.number;
        if (cellValue.includes(SYNTAX.VARIABLE_TABLE_SYNTAX)) return row.number - 1;
      }
    else return sheetOpts.beginTableAt;
    return undefined;
  }

  /** Get end table at by row */
  static endTableAt(row: exceljs.Row, sheetOpts: SheetExcelOption, isSampleExcel: boolean = true) {
    if (isSampleExcel)
      for (let i = 0; i < row.cellCount; i++) {
        const cellValue = row.getCell(i + 1) as unknown as string;
        if (typeof cellValue !== "string") continue;
        if (cellValue.includes(SYNTAX.END_TABLE_SYNYAX)) return row.number - 1;
        if (!cellValue.includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX) && cellValue.includes(SYNTAX.VARIABLE_TABLE_SYNTAX))
          return row.number - 1;
      }
    else if (!(row.values as any[])[sheetOpts.keyTableAt]) return row.number - 1;
    return undefined;
  }

  /** Get key of table by row */
  static columnTableIndex(row: exceljs.Row, sheetOpts: SheetExcelOption, isSampleExcel: boolean = true) {
    if (isSampleExcel)
      for (let i = 0; i < row.cellCount; i++) {
        const cellValue = row.getCell(i + 1) as unknown as string;
        if (typeof cellValue !== "string") continue;
        if ((cellValue + "").includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX)) return row.number + 1;
      }
    else return sheetOpts.keyTableAt;
    return undefined;
  }

  static isVariable(cellValue: any) {
    cellValue = cellValue + "";
    if (cellValue.includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX)) return false;
    return cellValue.includes(SYNTAX.VARIABLE_SYNTAX);
  }

  static getLabel(label: string, isVariable: boolean) {
    if (isVariable) return undefined;
    if (label.includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX)) return label.split("->").pop();
    return label;
  }

  static getAddress(address: string, section: SheetSection, isVariable: boolean) {
    return !isVariable ? address : section !== "header" ? address.split(/\d+/)[0] : address;
  }

  static getVariableValue(cell: exceljs.Cell): CellDataHelper["variableValue"] {
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
}
