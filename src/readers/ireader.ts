import { RawRecordType } from "../source/types/type.js";
import {Readable} from "stream";

export type ReaderResult = {type: RawRecordType} & Record<string, any>;
export type AddCallback = (data: ReaderResult|null) => any;

export abstract class IReader {
    protected readable: Readable = new Readable({objectMode: true});
    async open() {}
    async close() {}
    stream(): Readable {return this.readable}
    cancel() {}
    abstract get(add: AddCallback): Promise<void>;

    async start() {
        await this.get((data: ReaderResult|null) => this.readable.push(data));
    }
}