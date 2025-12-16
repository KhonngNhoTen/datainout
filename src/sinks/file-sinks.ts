import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { Sink } from "./base-sinks.js";

export type FileOutputType = "stream"| "file"|"buffer";
export enum FileExtension {
    EXCEL = "xlsx",
    CSV = "csv",
    JSON = "json",
    XML = "xml",
    HTML="html",
    PDF = "pdf"
}

export abstract class FileSink implements Sink {
    protected typeOutput!: FileOutputType;
    protected extension: FileExtension = FileExtension.EXCEL;
    abstract handle(chunk: MappedRecord[]): Promise<void>;
    
}