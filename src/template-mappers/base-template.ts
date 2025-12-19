import { RawRecordType, SourceType } from "../source/types/type.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { TableScope, Template } from "./types/template.type.js";
import fs from "fs";
export class BaseTemplate<T extends Template> {
    protected template!: T;
    protected cachedByName: Record<string, any> = {};

    constructor(template?: T) ;
    constructor(source: SourceType, rawType: RawRecordType); 
    constructor(arg1: any, arg2?: any) {
        if (arg2) {
            const template = arg1 as T;
            this.template = template;
            template?.fields.forEach(e => this.cachedByName[e.name] = e);
            template?.metadata.forEach(e => this.cachedByName[e.name] = e);
        } else {
            this.template = {
                fields: [],
                metadata: [],
                id: 0,
                number: 0,
                sourceType: arg1 as SourceType,
                type: arg2 as RawRecordType
            } as unknown as T
        }
    }

    private updateCache(name: string, field: any, action: "add"|"update"|"remove"="add") {
        if (action === "add") this.cachedByName[name] = field;
        else if (action === "update") this.cachedByName[name] = field;  
        else delete this.cachedByName[name];
    }

    add(field: T['fields'][0]) {
        if (field.scope === "table") this.template.fields.push(field);
        else this.template.metadata.push(field);

        this.updateCache(field.name, field, "add");
        return this;
    }

    remove(name: string, scope: TableScope = "table") {
        if (scope === "table") this.template.fields = this.template.fields.filter(field => field.name !== name);
        else if (scope === "metadata") this.template.metadata = this.template.metadata.filter(field => field.name !== name);
        
        this.updateCache(name, undefined, "remove");        
        return this;
    }

    update(name: string, field: T['fields'][0]) {
        const index = this.getIndexByName(name, field.scope);
        if (index < 0) return this;
        if (field.scope === "table") this.template.fields[index] = field;
        else if (field.scope === "metadata") this.template.metadata[index] = field;

        this.updateCache(name, field, "update");
        return this;
    }


    getByName(name: string, scope: TableScope = "table"): T['fields'][0] | undefined {
        // if (scope === "table")
        //     return this.template.fields.filter(field => field.name === name)?.[0] ?? undefined;
        // return this.template.metadata.filter(field => field.name === name)?.[0] ?? undefined;
        return this.cachedByName[name];
    }

    getIndexByName(name: string, scope: TableScope = "table"): number {
        if (scope === "table")
            return this.template.fields.findIndex(field => field.name === name);
        return this.template.metadata.findIndex(field => field.name === name);
    }

    getStructure(): T {
        return this.template;
    }


    save(fielPath: string) {
        fs.writeFileSync(fielPath, JSON.stringify(this.template, null, 2));
    }

}