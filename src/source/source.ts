import { ExcelHelper } from "../helpers/table.helper.js";
import { IReader } from "../readers/ireader.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { TableScope, Template } from "../template-mappers/types/template.type.js";
import { BaseSource } from "./base-source.js";
import { RawRecord } from "./types/type.js";



export class SourceAdapter extends BaseSource{
    async getIterator(): Promise<AsyncIterableIterator<RawRecord[]>> {
        const that = this;
        const template: Template = this.template;
        let metadata: any[] = [];
        let rawRecords: RawRecord[] = [];

        async function* iteractor() {
            for await (const record of that.reader.getIterator()) {
                let scope: TableScope = "table";
                if (record.type === "table")
                    scope = ExcelHelper.getScope(record.number, (template as TableTemplateOpts).table);
                if (scope === "metadata")
                    metadata.push(record);
                else {
                    const rawRecord: RawRecord = {
                        type: record.type,
                        metadata: metadata,
                        fields: record
                    };
                    rawRecords.push(rawRecord);
                }

                if (rawRecords.length === that.batchSize) {
                    yield rawRecords;
                    rawRecords = [];
                }
            }
            if (rawRecords.length > 0) yield rawRecords;
        }
        return iteractor();
    }
}
