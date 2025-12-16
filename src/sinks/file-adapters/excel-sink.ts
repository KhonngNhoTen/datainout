import { MappedRecord } from "../../transformer/types/transformer-dto.js";
import { FileSink } from "../file-sinks.js";

export class ExcelSink extends FileSink{
    async handle(chunk: MappedRecord[]) {
        chunk[0].metadata
    }
}