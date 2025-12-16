import { Events } from "../common/constant/events.js";
import { Engine } from "../engines/engine.js";
import { Sink } from "../sinks/base-sinks.js";
import { BaseSource } from "../source/base-source.js";
import { RawRecord } from "../source/types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { Template } from "../template-mappers/types/template.type.js";
import { BaseTransformer } from "../transformer/base-transformer.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { BaseValidator } from "../validators/base-validator.js";

export class Importer<T extends Template> extends Engine<T> {
    constructor(
        source: BaseSource,
        template: BaseTemplate<T>,
        transformer: BaseTransformer<Template>,
        validator: BaseValidator,
        sink: Sink,
        options: any,
    ) {
        super(source, template, transformer, validator, sink, options);
    }

    protected handleError() {

    }

    protected handleRecord(record: RawRecord): MappedRecord | undefined {
        const dto = record.type !== "object" ? this.transformer.parse(record) : record as any;
        const validate = this.validator.check(dto);
        if (validate.length > 0) {
            this.handleError;
            this.eventBus.emit(Events.onRecordError, { errors: validate });
            return undefined;
        }
        return dto;

    }
} 