import { RawRecordType } from "../source/types/type.js";
import {PassThrough} from "stream";
import { TableScope } from "../template-mappers/types/template.type.js";

export type ReaderResult = {type: RawRecordType, scope: TableScope} & Record<string, any>;
export type AddCallback = (data: ReaderResult|null) => any;

export abstract class IReader {
    protected readable: PassThrough = new PassThrough({objectMode: true});
    
    async open(opts?: any) {}

    async close() {}
    stream(): PassThrough {return this.readable}
    cancel() {}
    abstract get(add: AddCallback): Promise<void>;

    async start() {
        await this.get((data: ReaderResult|null) => this.readable.push(data));
    }
}