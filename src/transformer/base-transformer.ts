import { RawRecord } from "../source/types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { Template } from "../template-mappers/types/template.type.js";
import { MappedRecord } from "./types/transformer-dto.js";

export abstract class BaseTransformer<T extends Template> {
    protected templateStrct!: T;
    protected savedMetadata: Record<string, any> = {};
    
    constructor(template: BaseTemplate<T>) {
        this.templateStrct = template.getStructure();
    }

    abstract parse(record: RawRecord): MappedRecord;
}