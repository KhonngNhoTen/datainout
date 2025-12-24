import { RowModel, XlsxReadOptions, Style, Cell, Row } from "exceljs";
import { TemplateCellOpts } from "../template-mappers/types/table-template.type.js"
import { RawRecordType } from "../source/types/type.js";
import { TableScope } from "../template-mappers/types/template.type.js";

export type TableRowRaw = {
    cells: (Cell & {value: any})[],
    model: Partial<RowModel> | null;
    height: number;
    outlineLevel?: number;
    cellCount: number;
	actualCellCount: number;
    // name?: string;
    type: RawRecordType;
    id: string;
    number: number;
    scope: TableScope;
} & Partial<Style>;

export type _TableCellRaw = {
    value: any;
    formula?: string
} & Partial<Style>;

export type _TableRowRaw = {
    cells: _TableCellRaw[],
    height: number;
	hidden: boolean;
    outlineLevel?: number;
    model: Partial<RowModel> | null;
    type: RawRecordType;
    number: number;
    scope: TableScope;
} & Partial<Style>;

export type ExcelReaderConfig = XlsxReadOptions;

export type ReaderOpenOpts = {
    buffer?: Buffer;
    filePath?: string;
};