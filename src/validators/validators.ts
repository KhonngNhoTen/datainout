import { Template } from "ejs";
import { BaseValidator,  ValidationOptions } from "./base-validator.js";
import { ValidateResult } from "./types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { TableScope, TemplateField } from "../template-mappers/types/template.type.js";
import { MappedRecord } from "../transformer/types/transformer-dto.js";

export class Validators extends BaseValidator {
    checkField(value: any, fieldTemplate: TemplateField): ValidateResult {
        throw new Error("Method not implemented.");
    }
  
    constructor(template: BaseTemplate<any>, options?: ValidationOptions) {
        super(template, options);
    }

}