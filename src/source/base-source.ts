import { BaseTemplate } from "../template-mappers/base-template.js";
import { IReader } from "../readers/ireader.js";
import { Template } from "../template-mappers/types/template.type.js";
import { ExcelReader } from "../readers/excel-readers.js";
import { Readable, Writable } from "stream";
import { RawRecord } from "./types/type.js";
import { ReaderOpenOpts } from "../readers/type.js";

export type BaseSourceOptions = {
    batchSize?: number;
    concurrency?: number;
} & ReaderOpenOpts;

const IS_BASE_SOURCE = '$$_BaseSource';
export function isBaseSource (source: any): source is BaseSource {
    return source?.[IS_BASE_SOURCE];
}

export abstract class BaseSource {
    [IS_BASE_SOURCE]: boolean = true
    protected batchSize: number = 100;
    protected concurrency: number = 1;
    protected template!: Template;
    protected reader!: IReader;
    protected file?: string|Buffer|Readable;

    init(template: Template, options: BaseSourceOptions = {}) {
        this.template = template;
        this.reader = this.createReader();
        this.batchSize = options.batchSize ?? this.batchSize;
        this.concurrency = options.concurrency ?? this.concurrency;
    }

    createReader(): IReader {
        if (this.template.sourceType === "excel")
            return new ExcelReader()
        return new ExcelReader()
    }

    async open() {
        await this.reader.open({
            buffer: this.file instanceof Buffer ? this.file : undefined,
            filePath: typeof this.file === "string" ? this.file : undefined
        });
    }

    async close() {
        this.reader.close();
    }

    abstract getIterator(): Promise<AsyncIterableIterator<RawRecord[]>>;
}