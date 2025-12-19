import { ValidateResult } from "./types/type.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";
import { Template, TemplateField } from "../template-mappers/types/template.type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";

export type ValidationOptions = {
    errorStrategy?: 'fail-fast' | 'skip' | 'collect' | 'redirect-file';
}

export abstract class BaseValidator {
    protected templateStrct!: Template;
    protected template!: BaseTemplate<any>;
    protected options!: ValidationOptions;
    protected checkedMetadata: boolean = false;
    protected fieldNames: string[] = [];

    constructor(template: BaseTemplate<any>, options?: ValidationOptions) {
        this.template = template;
        this.templateStrct = template.getStructure();
        this.options = options ?? {
            errorStrategy: 'fail-fast'
        };
        this.fieldNames = this.templateStrct.fields.map(e => e.name);
    };

    check(dto: MappedRecord): ValidateResult[] {
        const validationResults: ValidateResult[] = [];
        if (!this.checkedMetadata)
            validationResults.push(...this.checkMetadata(dto.metadata));

        const fields: any = dto.fields;
        for (let j = 0; j < this.fieldNames.length; j++) {
            const name = this.fieldNames[j];
            const templateStrctField = this.template.getByName(name);
            const validate = this.checkField(fields[name], templateStrctField);
            if (validate.validate === false) validationResults.push(validate);
        }

        return validationResults;
    }

    private checkMetadata(metadata: Record<string, any>): ValidateResult[] {
        const validationResults: ValidateResult[] = [];
        const keys = metadata ? Object.keys(metadata) : [];
        for (let i = 0; i < keys.length; i++) {
            const key = keys[i];
            const field = metadata[key];
            const templateStrctField = this.template.getByName(key, "metadata");
            if (!templateStrctField) continue;
            const validate = this.checkField(field, templateStrctField);
            if (validate.validate === false) validationResults.push(validate);
        }

        this.checkedMetadata = true;
        return validationResults
    }

    abstract checkField(value: any, fieldTemplate: TemplateField): ValidateResult;
    
}