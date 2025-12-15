import { RowModel, XlsxReadOptions } from "exceljs";
import { TemplateCellOpts } from "../template-mappers/types/table-template.type.js"
import { RawRecordType } from "../source/types/type.js";

export type TableRowRaw = {
    cells: TemplateCellOpts[],
    model: Partial<RowModel> | null;
    height: number;
    outlineLevel?: number;
    cellCount: number;
	actualCellCount: number;
    name?: string;
    type: RawRecordType;
    id: string;
    number: number;
}

export type ExcelReaderConfig = XlsxReadOptions;

export type ReaderOpenOpts = {
    buffer?: Buffer;
    filePath?: string;
};