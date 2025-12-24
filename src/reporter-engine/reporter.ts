import { BaseTemplate } from "../template-mappers/base-template.js";
import { BaseValidator } from "../validators/base-validator.js";
import { Template } from "../template-mappers/types/template.type.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { Engine } from "../engines/engine.js";
import { Events } from "../common/constant/events.js";
import { BaseSource } from "../source/base-source.js";
import { FileOutputType, ExportSink } from "../sinks/export-sinks.js";

export class Reporter<T extends Template, U extends FileOutputType> extends Engine<T> {
    declare sink: ExportSink<U>;
    constructor(
        source: BaseSource,
        template: BaseTemplate<T>,
        validator: BaseValidator,
        sink: ExportSink<U>,
    ) {
        super({source, template, validator, sink});
    }

    protected handleError() {}

    protected handleRecord(dto: MappedRecord): MappedRecord | undefined {
        const validate = this.validator.check(dto);
        if (validate.length > 0) {
            this.handleError;
            this.eventBus.emit(Events.onRecordError, { errors: validate });
            return undefined;
        }
        return dto;

    }

    async export(): Promise<Awaited<U extends "buffer" ? Promise<Buffer> : U extends "file" ? Promise<void> : any>> {
        await this.open();

        await this.source.start();

        await this.engineStream.waitingDone();

        return await this.sink.export() as any;
    }
}