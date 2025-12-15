import { Cell, Row, RowModel } from "exceljs";
import { Template, TemplateField } from "./template.type.js";

export type Style = {
    numFmt?: string;
    font?: Cell["font"];
    alignment?: Cell['alignment'];
    protection?: Cell['protection'];
    border?: Cell['border'];
    fill?: Cell['fill'];
    height?: number;
}
export type TemplateCellOpts = {
    isVariable: boolean,
    address: {
        row: number,
        column: string
    },
    formula: Cell["formula"],
    
} & TemplateField & Style

export type TableTemplateOpts = {
    table: {
        startAt: number,
        endAt?: number
    },
    fields: TemplateCellOpts[],
    // model: Partial<RowModel> | null;
    // height: number;
    // outlineLevel?: number;
    // cellCount: number;
	// actualCellCount: number;
} & Omit<Template, "fields"> & Style;

