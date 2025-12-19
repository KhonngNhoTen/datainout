import { BaseTemplate } from "../template-mappers/base-template.js";
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
    protected template!: BaseTemplate<any>;
    protected filePath!: string;

    constructor(
        typeOutput: FileOutputType,
        extension: FileExtension = FileExtension.EXCEL,
        template: BaseTemplate<any>
    ) {
        this.typeOutput = typeOutput;
        this.extension = extension;
        this.template = template;

    }

    abstract handle(chunk: MappedRecord[]): Promise<void>;
 
}