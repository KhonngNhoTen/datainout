import { Events } from "../common/constant/events.js";
import { EventBus } from "../common/event-bus.js";
import { Sink } from "../sinks/base-sinks.js";
import { BaseSource } from "../source/base-source.js";
import { RawRecord } from "../source/types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { Template } from "../template-mappers/types/template.type.js";
import { BaseTransformer } from "../transformer/base-transformer.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { BaseValidator } from "../validators/base-validator.js";
import { Writable } from "stream";

type EngineContructorOptions<T extends Template> = {
    source?: BaseSource,
    template?: BaseTemplate<T>,
    transformer?: BaseTransformer<Template>,
    validator?: BaseValidator,
    sink?: Sink,
}
export abstract class Engine<T extends Template> {
    protected source!: BaseSource;
    private template!: BaseTemplate<T>;
    protected transformer!: BaseTransformer<Template>
    protected validator!: BaseValidator;
    protected sink!: Sink;
    protected engineStream: EngineStream = new EngineStream({
        onChunk: (chunk) => this.handle(chunk),
        onEnd: () => this.close(),
    });
    protected eventBus: EventBus = new EventBus();

    constructor(
        opts: EngineContructorOptions<T>
    ) {
        this.source = opts.source as any;
        this.template = opts.template as any;
        this.transformer = opts.transformer as any;
        this.validator = opts.validator as any;
        this.sink = opts.sink as any;
        this.source.stream().pipe(this.engineStream);
    }

  
    protected async open() {
        this.eventBus?.emit(Events.onFile);
        await this.source.open();
        await this.sink.open?.();
    }

    protected async close() {
        await this.source.close();
        await this.sink.close?.();
        this.eventBus?.emit(Events.finishedFile);
    }

    protected async handle(records: RawRecord[]) {
        const data: any[] = []
        for (const record of records) {
            this.eventBus?.emit(Events.onRecord);

            let dto: MappedRecord | undefined = 
                record.type === "object" ? 
                record as any : 
                this.transformer.parse(record);

            if (dto === undefined) continue;
            dto = this.handleRecord(dto);
            data.push(dto);
            this.eventBus?.emit(Events.finishedRecord);

        }

        await this.sink.handle(data);

    }

    protected handleError() { }

    protected handleRecord(dto: MappedRecord): MappedRecord | undefined { return undefined }

    get Template(): BaseTemplate<T> { return this.template; }

    public on(event: `${Events}`, listener: (data: any) => Promise<void>) { this.eventBus.on(event, listener); }

    public off(event: `${Events}`) { this.eventBus.off(event); }

}

export class EngineStream extends Writable {
    private onChunk: (chunk: any) => Promise<void>;
    private onEnd?: () => Promise<void>;
    // onStart?: () => void;
    private callBackDone?: (value?: any) => void;

    constructor(
        opts: {
            onChunk: (chunk: any) => Promise<void>,
            onEnd?: () => Promise<void>
            // onStart?: () => void,
        }
    ) {
        super({ objectMode: true });
        this.onChunk = opts.onChunk;
        this.onEnd = opts.onEnd;
        // this.onStart = opts.onStart
    }

    waitingDone () {
        return new Promise((r) => { this.callBackDone = r; });
    }

    _write(
        chunk: any,
        encoding: BufferEncoding,
        callback: (error?: Error | null) => void
    ): void {
        this.onChunk(chunk)
            .then(() => callback())
            .catch((err) => {
                callback(err);
            });
    }

    _final(callback: (error?: Error | null) => void) {
        this.onEnd?.().
            then(() => {
                callback();
                this.callBackDone?.();
            }).
            catch((e) => callback(e as Error));
    }
}