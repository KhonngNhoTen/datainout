import { RawRecord } from "../source/types/type.js";
// import { TableTemplate } from "../template-mappers/table-template";
import { TableTemplateOpts, TemplateCellOpts } from "../template-mappers/types/table-template.type.js";
import { BaseTransformer } from "./base-transformer.js";
import exceljs, { Row } from "exceljs";
import { MappedRecord } from "./types/transformer-dto.js";
import { ExcelHelper } from "../helpers/table.helper.js";
import { TableScope, TemplateField } from "../template-mappers/types/template.type.js";
import { TableRowRaw } from "../readers/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";

export class TableTransformer extends BaseTransformer<TableTemplateOpts> {

    private groupByAddress: Record<string, TemplateField> = {};

    constructor(template: BaseTemplate<TableTemplateOpts>) {
        super(template);
        this.templateStrct.fields.forEach(e => this.groupByAddress[`${e.scope}:${e.address}`] = e);
        this.templateStrct.metadata.forEach(e => this.groupByAddress[`${e.scope}:${e.address}`] = e);
    }

    parse(record: RawRecord) {
        const table: MappedRecord = {} as any;
        const _record = record as unknown as TableRowRaw;
        const scope = ExcelHelper.getScope(_record.number, this.templateStrct.table);
        let row: any = {};

        _record.cells.forEach(cell => {
            const value = this.parseCell(cell, scope);
            if (!value) return;
            if (scope === "table") row[value.name] = cell.value;
            else if (scope === "metadata") this.savedMetadata[value.name] = value;
        });
        table.fields.push(row);
        table.type = record.type;
        table.metadata = this.savedMetadata;
        return table;
    }

    parseCell(cell: TemplateCellOpts, scope: TableScope) {
        if (!this.groupByAddress[`${scope}:${cell.address}`]) return undefined;
        const field = this.groupByAddress[`${scope}:${cell.address}`];
        return {
            name: field.name,
            value: field
        }
    }
}