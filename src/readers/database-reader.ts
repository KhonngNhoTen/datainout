import { IReader, ReaderResult } from "./ireader.js";

export abstract class DatabaseReader implements IReader {
    abstract getIterator(): AsyncIterableIterator<ReaderResult>;
    
    cancel(): void {}
    async close() {}
    async open() {}

}