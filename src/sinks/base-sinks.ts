import { MappedRecord } from "../transformer/types/transformer-dto.js";

export interface Sink {
    open?(): Promise<void>;
    handle(chunk: MappedRecord[]): Promise<void>;
    close?(): Promise<void>;
}