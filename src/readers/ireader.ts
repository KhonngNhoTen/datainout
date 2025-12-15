import { RawRecordType } from "../source/types/type.js";
import { ReaderOpenOpts } from "./type.js";

export type ReaderResult = {type: RawRecordType} & Record<string, any>
export interface IReader {
    open(options: ReaderOpenOpts): Promise<void>;
    close(): Promise<void>;
    getIterator(): AsyncIterableIterator<ReaderResult>;
    cancel(): void;
}