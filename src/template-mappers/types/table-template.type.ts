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
    address: string;
    addressDetail: {
        row: number;
        column: string;
    };
    formula: Cell["formula"],

} & TemplateField & Style

export type TableTemplateOpts = Omit<Template, "fields"|"metadata"> & Style & {
    table: {
        startAt: number,
        endAt?: number
    },
    fields: TemplateCellOpts[],
    metadata: TemplateCellOpts[],
};

