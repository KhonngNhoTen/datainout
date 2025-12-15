import { SourceType } from "../source/types/type.js";
import { TableScope, Template } from "./types/template.type.js";

export class BaseTemplate<T extends Template> {
    protected template!: T;
    protected cachedByName: Record<string, any> = {};

    constructor(template?: T) {
        if (template) {
            this.template = template;
            template.fields.forEach(e => this.cachedByName[e.name] = e);
            template.metadata.forEach(e => this.cachedByName[e.name] = e);
        }
    }

    add(field: T['fields'][0], scope: TableScope = "table") {
        if (scope === "table") this.template.fields.push(field);
        else this.template.metadata.push(field);
        return this;
    }

    remove(name: string, scope: TableScope = "table") {
        if (scope === "table") this.template.fields = this.template.fields.filter(field => field.name !== name);
        else if (scope === "metadata") this.template.metadata = this.template.metadata.filter(field => field.name !== name);
        return this;
    }

    update(name: string, field: T['fields'][0], scope: TableScope = "table") {
        const index = this.getIndexByName(name, scope);
        if (index < 0) return this;
        if (scope === "table") this.template.fields[index] = field;
        else if (scope === "metadata") this.template.metadata[index] = field;
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

}