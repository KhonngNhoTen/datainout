import * as exceljs from "exceljs";
import { AttributeType, SheetExcelOption, SheetSection } from "./common-type.js";
import { ExcelTemplateManager } from "../core/Template.js";
import { CellImportOptions } from "./import-template.type.js";

export type CellDataHelper = {
  rowIndex: number;
  isVariable: boolean;
  label?: string;
  variableValue?: { fieldName: string; type: AttributeType };
  detail: exceljs.Cell;
  address: string;
  section: SheetSection;
  beginTableAt: number;
  endTableAt: number;
  formula?: exceljs.Cell["formula"];
  //   rowCount: number;
};

export type RowDataHelper = {
  rowIndex: number;
  detail: exceljs.Row;
  section: SheetSection;
  cells: CellDataHelper[];
  beginTableAt?: number;
  endTableAt?: number;
};

export type SheetDataHelper = {
  sheetIndex: number;
  columnIndex: number;
  name: string;
  detail: exceljs.Worksheet;
  beginTableAt: number;
  endTableAt: number;
  rowCount: number;
  lastestRow: RowDataHelper;
  firstRow: RowDataHelper;
};

export type ExcelReaderHelperOptions = {
  onSheet?: (sheet: SheetDataHelper) => Promise<any>;
  onRow?: (row: RowDataHelper) => Promise<any>;
  onCell?: (cell: CellDataHelper) => Promise<any>;
  isSampleExcel?: boolean;
  template?: SheetExcelOption<any>[];
  templateManager: ExcelTemplateManager<CellImportOptions>;
};
