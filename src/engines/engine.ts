import { Events } from "../common/constant/events.js";
import { EventBus } from "../common/event-bus.js";
import { Sink } from "../sinks/base-sinks.js";
import { SourceAdapter } from "../source/source.js";
import { RawRecord } from "../source/types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { Template } from "../template-mappers/types/template.type.js";
import { BaseTransformer } from "../transformer/base-transformer.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { BaseValidator } from "../validators/base-validator.js";

export abstract class Engine<T extends Template> {
    protected source!: SourceAdapter;
    private template!: BaseTemplate<T>;
    protected transformer!: BaseTransformer<Template>
    protected validator!: BaseValidator;
    protected sink!: Sink;
    protected options!: any;
    protected eventBus!: EventBus;

    constructor(
        source: SourceAdapter,
        template: BaseTemplate<T>,
        transformer: BaseTransformer<Template>,
        validator: BaseValidator,
        sink: Sink,
        options: any,
    ) {
        this.options = options;
        this.source = source;
        this.template = template;
        this.transformer = transformer;
        this.validator = validator;
        this.sink = sink
    }

    async run() {
        this.eventBus.emit(Events.onFile, {});
        await this.source.open();
        await this.sink.open?.();
        let data: MappedRecord[] = [];
        for await (const records of await this.source.getIterator()) {
            data = []
            for await (const record of records) {
                this.eventBus.emit(Events.onRecord, {});
                const dto = this.handleRecord(record);
                if (!dto) continue;
                data.push(dto)
                this.eventBus.emit(Events.finishedRecord, {});
            }
            await this.sink.handle(data);
        }
        await this.sink.close?.();
        await this.source.close();
        this.eventBus.emit(Events.finishedFile, {});
    }

    get Template(): BaseTemplate<T> {
        return this.template;
    }

    protected handleError() { }

    protected handleRecord(record: RawRecord) {
        const dto = this.transformer.parse(record);
        const validate = this.validator.check(dto);
        if (validate.length > 0) {
            this.handleError;
            this.eventBus.emit(Events.onRecordError, { errors: validate });
            return undefined;
        }
        return dto;
    }

    public on(event: Events, listener: (data: any) => Promise<void>) {
        this.eventBus.on(event, listener);
    }
    
    public off(event: Events) {
        this.eventBus.off(event);
    }
}