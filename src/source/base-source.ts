import { BaseTemplate } from "../template-mappers/base-template.js";
// import { AddCallback, IReader } from "../readers/ireader.js";
import { Template } from "../template-mappers/types/template.type.js";
import { ExcelReader } from "../readers/excel-readers.js";
import { PassThrough, Readable, Writable } from "stream";
import { RawRecord } from "./types/type.js";
import { ReaderOpenOpts } from "../readers/type.js";

export type BaseSourceOptions = {
    batchSize?: number;
    concurrency?: number;
} & ReaderOpenOpts;

const IS_BASE_SOURCE = '$$_BaseSource';
export function isBaseSource(source: any): source is BaseSource {
    return source?.[IS_BASE_SOURCE];
}


export type CallbackSource = (data: RawRecord[]|null) => any;
export abstract class BaseSource {
    [IS_BASE_SOURCE]: boolean = true
    protected batchSize: number = 100;
    protected concurrency: number = 1;
    protected readable: PassThrough = new PassThrough({ objectMode: true });

    constructor(batchSize?: number, concurrency?: number) {
        this.batchSize = batchSize ?? this.batchSize;
        this.concurrency = concurrency ?? this.concurrency;
    }

    stream() { return this.readable; }
    async open() { }
    async close() { }
   
    abstract get(add: CallbackSource): Promise<void>;

    async start() {
        await this.get((data) => this.readable.push(data));
    }
}