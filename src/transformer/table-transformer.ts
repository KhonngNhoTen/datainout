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
    private cachedMetadata: any = {};
    private visitedMetdata = false;

    constructor(template: BaseTemplate<TableTemplateOpts>) {
        super(template);
    }


    parse(record: RawRecord): MappedRecord {
        this.setMetadata(record);

        const dto: any = {};
        const _record = record as unknown as TableRowRaw;
        const groupValues = this.groupByAddress(_record.cells, "table");

        this.templateStrct.fields.forEach(f => {
            const cell = groupValues[f.address];
            if(!cell) return;
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
        if(this.visitedMetdata) return;
        if(this.templateStrct.metadata?.length <= 0) return;

        const metadata: any = {};
        const cells: (TemplateCellOpts & {value: any})[] = [];
        (record.metadata as TableRowRaw[]).forEach(m => cells.push(...m.cells));
        const groupValues = this.groupByAddress(cells, "metadata");

        this.templateStrct.metadata.forEach(m => {
            const cell = groupValues[m.address];
            if(!cell) return;
            
            const value = m.setValue ? m.setValue(cell.value) : cell.value;
            metadata[m.name] = value;
        });

        this.cachedMetadata = metadata;
        this.visitedMetdata = true;
    }

    groupByAddress(cells: (TemplateCellOpts & {value: any})[], scope: TableScope) {
        if(scope === "metadata") {
            return cells.reduce((acc: any, c) => {
                acc[c.address] = c;
                return acc;
            }, {})
        } else {
            return cells.reduce((acc: any, c) => {
                acc[c.addressDetail.column] = c;
                return acc;
            }, {})
        }
    }
}