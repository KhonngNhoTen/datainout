import { RawRecord } from "../source/types/type.js";
import { Template } from "../template-mappers/types/template.type.js";
import { BaseTransformer } from "./base-transformer.js";
import { MappedRecord } from "./types/transformer-dto.js";

export class EventTransformer extends BaseTransformer<Template> {
    parse(record: RawRecord): MappedRecord {
        return {
            type: "table",
            fields: record.fields.reduce((acc: any, f: any) => ({ ...acc, [f.name]: f }), {}),
            metadata: record.metadata?.reduce((acc: any, f: any) => ({ ...acc, [f.name]: f }), {}) ?? {}
        }
        // const fieldNames = Object.keys(record.fields);
        // const dto: MappedRecord = {type: "event","fields": {}, metadata: {}};
        // for (let i = 0; i < fieldNames.length; i++) {
        //     const name = fieldNames[i];
        //     const templateField = this.template.getByName(name);
        //     dto.fields[name] = templateField ? this.parseField(record.fields[name], templateField): record.fields[name];
        // }

        // return dto;
    }

}