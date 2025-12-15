import { ExcelHelper } from "../helpers/table.helper.js";
import { ExcelReader } from "../readers/excel-readers.js";
import { IReader } from "../readers/ireader.js";
import { ReaderOpenOpts } from "../readers/type.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { TableScope, Template } from "../template-mappers/types/template.type.js";
import { RawRecord } from "./types/type.js";

export type BaseSourceOptions = {
    batchSize?: number;
    concurrency?: number;
} & ReaderOpenOpts;

export class SourceAdapter {
    protected options: Partial<BaseSourceOptions>;
    protected reader!: IReader;
    protected template!: Template

    constructor(template: Template, options: BaseSourceOptions = {}) {
        this.options = {
            batchSize: 100,
            concurrency: 1,
            buffer: options.buffer,
            filePath: options.filePath
        };
        this.template = template;

        this.reader = this.createReader();
    }

    async open(): Promise<void> {
        await this.reader.open({
            buffer: this.options.buffer,
            filePath: this.options.filePath
        });
    }
    async close(): Promise<void> {

    }

    createReader(): IReader {
        if(this.template.sourceType === "excel")
            return new ExcelReader()
        return new ExcelReader()
    }

    async getIterator(): Promise<AsyncIterableIterator<RawRecord[]>> {
        const reader = this.reader;
        const options = this.options;
        const template: Template = this.template;
        let metadata: any[] = [];
        let rawRecords: RawRecord[] = [];

        async function* iteractor() {
            for await (const record of reader.getIterator()) {
                let scope: TableScope = "table";
                if(record.type === "table") 
                    scope = ExcelHelper.getScope(record.number, (template as TableTemplateOpts).table);
                if(scope === "metadata")
                    metadata.push(record);
                else {
                    const rawRecord: RawRecord = {
                        type: record.type,
                        metadata: metadata,
                        fields: record
                    };
                    rawRecords.push(rawRecord);
                }

                if(rawRecords.length === options.batchSize) {
                    yield rawRecords;
                    rawRecords = [];
               }
            }
            if(rawRecords.length > 0) yield rawRecords;
        }
        return iteractor();
    }

    public get Options() {
        return this.options;
    }
}
