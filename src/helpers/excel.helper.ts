import { AttributeType, SheetExcelOption, SheetSection } from "../common/types/common-type.js";
import * as exceljs from "exceljs";
import { CellDataHelper, ExcelReaderHelperOptions, RowDataHelper, SheetDataHelper } from "../common/types/excel-reader-helper.type.js";
import { CellImportOptions } from "../common/types/import-template.type.js";
import { ExcelTemplateManager } from "../common/core/Template.js";

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
  // private template: SheetExcelOption<CellImportOptions> = [];
  private templateManager: ExcelTemplateManager<CellImportOptions>;
  private onSheet?: (row: SheetDataHelper) => Promise<any>;
  private onRow?: (row: RowDataHelper) => Promise<any>;
  private onCell?: (cell: CellDataHelper) => Promise<any>;

  isStop: boolean = false;

  constructor(opts: ExcelReaderHelperOptions) {
    this.onCell = opts?.onCell;
    this.onRow = opts?.onRow;
    this.onSheet = opts?.onSheet;
    this.isSampleExcel = opts?.isSampleExcel ?? true;
    this.templateManager = opts.templateManager;
  }

  async load(file: string): Promise<any>;
  async load(buffer: Buffer): Promise<any>;
  async load(arg: unknown) {
    const workBook = new exceljs.Workbook();
    if (arg instanceof Buffer) await workBook.xlsx.load(arg as unknown as exceljs.Buffer);
    else await workBook.xlsx.readFile(arg as string);

    for (let i = 0; !this.isStop && i < workBook.worksheets.length; i++) {
      const workSheet = workBook.getWorksheet(i + 1);
      if (workSheet) await this.eachSheet(workSheet, i + 1);
    }
  }

  private async eachSheet(sheet: exceljs.Worksheet, sheetIndex: number) {
    this.templateManager.SheetIndex = sheetIndex - 1;
    const sheetDesc = this.templateManager.SheetTemplate;
    const trackingRows: RowDataHelper[] = [];

    let columnIndex = DEFAULT_COLUMN_INDEX;
    let beginTable = DEFAULT_BEGIN_TABLE;
    let endTable = DEFAULT_END_TABLE;
    for (let i = 1; !this.isStop && i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      this.templateManager.defineActualTableStartRow(ReaderExceljsHelper.beginTableAt(row, sheetDesc, this.isSampleExcel));
      this.templateManager.defineActualTableStartRow(ReaderExceljsHelper.endTableAt(row, sheetDesc, this.isSampleExcel));
      columnIndex = ReaderExceljsHelper.columnTableIndex(row, sheetDesc, this.isSampleExcel) ?? columnIndex;

      endTable = this.templateManager.ActualTableEndRow ?? DEFAULT_END_TABLE;
      beginTable = this.templateManager.ActualTableStartRow ?? DEFAULT_BEGIN_TABLE;

      const section = ReaderExceljsHelper.getSection(row, beginTable, endTable);

      const cells: CellDataHelper[] = [];
      for (let j = 0; j < row.cellCount; j++) {
        if (!row.getCell(j + 1)?.value) continue;
        const cell = this.convertCell(row.getCell(j + 1), section, beginTable, endTable);
        cells.push(cell);
        if (this.onCell) await this.onCell(cell);
      }

      if (this.onRow) {
        const rowDataHelper = this.convertRow(row, section, cells);
        if (i === 1) trackingRows.push(rowDataHelper);
        if (i === sheet.rowCount) trackingRows.push(rowDataHelper);
        await this.onRow(rowDataHelper);
      }
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
        lastestRow: trackingRows[1],
        firstRow: trackingRows[0],
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
      beginTableAt: cells[0]?.beginTableAt ?? DEFAULT_BEGIN_TABLE,
      endTableAt: cells[0]?.endTableAt ?? DEFAULT_END_TABLE,
    };
  }

  static getSection(row: exceljs.Row, beginTable: number, endTable: number): SheetSection;
  static getSection(rowIndex: number, beginTable: number, endTable: number): SheetSection;
  static getSection(row: unknown, beginTable: number, endTable: number): SheetSection {
    const rowIndex = typeof row === "number" ? row : (row as exceljs.Row).number;
    let section: SheetSection = "header";
    if (beginTable === DEFAULT_BEGIN_TABLE) section = "header";
    else if (beginTable !== DEFAULT_BEGIN_TABLE && rowIndex < beginTable) section = "header";
    else if (rowIndex > endTable && endTable !== DEFAULT_END_TABLE) section = "footer";
    else if (rowIndex > beginTable) section = "table";

    return section;
  }

  /** Get begin table at by row */
  static beginTableAt(row: exceljs.Row, sheetOpts?: SheetExcelOption<CellImportOptions>, isSampleExcel: boolean = true) {
    if (isSampleExcel)
      for (let i = 0; i < row.cellCount; i++) {
        const cellValue = row.getCell(i + 1).value as unknown as string;
        if (cellValue === undefined || typeof cellValue !== "string") continue;
        if (cellValue.includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX)) return row.number;
        if (cellValue.includes(SYNTAX.VARIABLE_TABLE_SYNTAX)) return row.number - 1;
      }
    else if (sheetOpts) return sheetOpts.beginTableAt;

    return undefined;
  }

  /** Get end table at by row */
  static endTableAt(row: exceljs.Row, sheetOpts?: SheetExcelOption<CellImportOptions>, isSampleExcel: boolean = true) {
    if (isSampleExcel)
      for (let i = 0; i < row.cellCount; i++) {
        const cellValue = row.getCell(i + 1).value as unknown as string;
        if (cellValue === undefined || typeof cellValue !== "string") continue;
        if (cellValue.includes(SYNTAX.END_TABLE_SYNYAX)) return row.number - 1;
        if (!cellValue.includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX) && cellValue.includes(SYNTAX.VARIABLE_TABLE_SYNTAX)) return row.number;
      }
    else if (sheetOpts) {
      const cellvalues = (row?.values as any[]) ?? [];
      if (cellvalues && !cellvalues[sheetOpts.keyTableAt] && row.number > sheetOpts.beginTableAt) return row.number - 1;
    }
    return undefined;
  }

  static isNullableRow(row: exceljs.Row) {
    return (row?.values as any[])?.reduce((acc, val) => acc && !!!val, true) ?? false;
  }

  /** Get key of table by row */
  static columnTableIndex(row: exceljs.Row, sheetOpts?: SheetExcelOption<CellImportOptions>, isSampleExcel: boolean = true) {
    if (isSampleExcel)
      for (let i = 0; i < row.cellCount; i++) {
        const cellValue = row.getCell(i + 1).value as unknown as string;
        if (cellValue === undefined || typeof cellValue !== "string") continue;
        if ((cellValue + "").includes(SYNTAX.INDEX_COLUMN_TABLE_SYNTAX)) return row.number + 1;
      }
    else if (sheetOpts) return sheetOpts.keyTableAt;
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

  static splitAddress(address: string) {
    const col = address.split(/\d+/)[0];
    const row = address.split(/[a-zA-Z]/)[1];
    return { col, row };
  }
}
