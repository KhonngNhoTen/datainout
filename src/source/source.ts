import { ExcelHelper } from "../helpers/table.helper.js";
import { AddCallback, IReader } from "../readers/ireader.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { TableScope, Template } from "../template-mappers/types/template.type.js";
import { BaseSource, CallbackSource } from "./base-source.js";
import { RawRecord } from "./types/type.js";
import {Readable} from "stream";


// export class SourceAdapter extends BaseSource {
//     protected reader!: IReader;
//     protected template!: Template;

//     constructor(template: Template, reader: IReader) {
//         super();
//         this.template = template;
//         this.reader = reader;
//     }

//     async getIterator(): Promise<AsyncIterableIterator<RawRecord[]>> {
//         const that = this;
//         const template: Template = this.template;
//         let metadata: any[] = [];
//         let rawRecords: RawRecord[] = [];

//         async function* iteractor() {
//             for await (const record of that.reader.getIterator()) {
//                 let scope: TableScope = "table";
//                 if (record.type === "table")
//                     scope = ExcelHelper.getScope(record.number, (template as TableTemplateOpts).table);
//                 if (scope === "metadata")
//                     metadata.push(record);
//                 else {
//                     const rawRecord: RawRecord = {
//                         type: record.type,
//                         metadata: metadata,
//                         fields: record
//                     };
//                     rawRecords.push(rawRecord);
//                 }

//                 if (rawRecords.length === that.batchSize) {
//                     yield rawRecords;
//                     rawRecords = [];
//                 }
//             }
//             if (rawRecords.length > 0) yield rawRecords;
//         }
//         return iteractor();
//     }


//     async open() {
//         await this.reader?.open();
//     }
    
//     async close() {
//         await this.reader?.close();
//     }
// }


export class Source extends BaseSource {
   
    protected reader!: IReader;
    protected template!: Template;
    protected readerStream!: Readable;

    constructor(template: Template, reader: IReader, batchSize?: number, concurrency?: number) {
        super(batchSize, concurrency);
        this.template = template;
        this.reader = reader;
        reader.stream().pipe(this.readable);
    }

    async get(add: CallbackSource): Promise<void> {}

    async start() {
        await this.reader.start();
    }
}