import { ExcelHelper } from "../helpers/table.helper.js";
import { AddCallback, IReader, ReaderResult } from "../readers/ireader.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { TableScope, Template } from "../template-mappers/types/template.type.js";
import { BaseSource, CallbackSource } from "./base-source.js";
import { RawRecord } from "./types/type.js";
import {PassThrough, Readable, Transform, TransformCallback} from "stream";


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

    constructor(template: Template, reader: IReader, batchSize?: number, concurrency?: number) {
        super(batchSize, concurrency);
        this.template = template;
        this.reader = reader;
        this.readable = new SourceTransform(this.batchSize);
        reader.stream().pipe(this.readable);
    }

    async open(): Promise<void> {
        await this.reader.open();
    }

    async close(): Promise<void> {
        await this.reader.close();
    }

    async get(add: CallbackSource): Promise<void> {}

    async start() {
        await this.reader.start();
    }
}

class SourceTransform extends Transform{
    private batchSize: number;
    private metadata: any[] = [];
    private rawRecords: RawRecord[] = [];
    private index: number = 0;


    constructor(batchSize: number) {
        super({objectMode: true});
        this.batchSize = batchSize;
    }

    _transform(data: ReaderResult, _: any, cb: any): void {
        if(data.scope === "metadata") this.metadata.push(data);
        else {
            this.rawRecords.push({
                type: data.type,
                metadata: this.metadata,
                fields: data,
            });

            this.index++;
            if(this.index === this.batchSize) {
                this.index = 0;
                this.push(this.rawRecords);
                this.rawRecords = [];
            }
        }
        cb();
    }

    _flush(cb: any) {
    if (this.rawRecords.length > 0) {
        this.push(this.rawRecords);
        this.rawRecords = [];
    }
    cb();
}
}