import { RawRecord, RawRecordType } from "../source/types/type.js";
// import { TableTemplate } from "../template-mappers/table-template";
import { TableTemplateOpts, TemplateCellOpts } from "../template-mappers/types/table-template.type.js";
import { BaseTransformer } from "./base-transformer.js";
import exceljs, { Cell, Row } from "exceljs";
import { MappedRecord } from "./types/transformer-dto.js";
import { ExcelHelper } from "../helpers/table.helper.js";
import { TableScope, TemplateField } from "../template-mappers/types/template.type.js";
import { TableRowRaw } from "../readers/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";

export class TableTransformer extends BaseTransformer<TableTemplateOpts> {
    private cachedMetadata: any = {};
    private visitedMetdata = false;

    constructor(template: BaseTemplate<TableTemplateOpts>) {
        super(template);
    }


    parse(record: RawRecord): MappedRecord {

        this.setMetadata(record);
        const dto: any = {};
        const _record = record as {
            type: RawRecordType;
            fields: TableRowRaw;
            metadata?: TableRowRaw[];
        };
        const groupValues = this.groupByAddress(_record.fields.cells, "fields");

        this.templateStrct.fields.forEach(f => {
            const cell = groupValues[f.addressDetail.columnIndex];
            if (!cell) return;
            const value = f.setValue ? f.setValue(cell.value) : cell.value;
            dto[f.name] = value;
        });

        return {
            fields: dto,
            metadata: this.cachedMetadata,
            type: "table"
        };
    }

    setMetadata(record: RawRecord) {
        if (this.visitedMetdata) return;
        if (this.templateStrct.metadata?.length <= 0) return;
        const metadata: any = {};
        const cells: (Cell & { value: any })[] = [];
        (record.metadata as TableRowRaw[]).forEach(m => cells.push(...m.cells));
        const groupValues = this.groupByAddress(cells, "metadata");
        this.templateStrct.metadata.forEach(m => {
            const cell = groupValues[m.addressDetail.column];
            if (!cell) return;

            const value = m.setValue ? m.setValue(cell.value) : cell.value;
            metadata[m.name] = value;
        });

        this.cachedMetadata = metadata;
        this.visitedMetdata = true;
    }

    groupByAddress(cells: (Cell & { value: any })[], scope: TableScope) {
        return cells.reduce((acc: any, c) => {
            const key = scope === "fields" ? c.col : c.address;
            acc[key] = c;
            return acc;
        }, {});
    }
}