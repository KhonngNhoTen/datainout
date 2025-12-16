import { BaseValidator, ValidationOptions } from "./base-validator.js";
import { ValidateResult } from "./types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { TemplateField } from "../template-mappers/types/template.type.js";

export class Validators extends BaseValidator {
    applyTemplate(value: any, fieldTemplate: TemplateField): { value: any, checked: ValidateResult } {
        let checked: ValidateResult = { validate: true };
        if (!fieldTemplate) checked = { validate: false, msg: "Missing template field" };

        if (fieldTemplate.setValue)
            value = fieldTemplate.setValue(value);

        if (fieldTemplate.required === true && value === undefined) checked = { validate: false, msg: "Field is required" };
        if (fieldTemplate.validate) {
            const rs = fieldTemplate.validate(value);
            if (rs.validate === false) checked = { validate: rs.validate, msg: rs.msg ?? "Validate fail" };
        }

        return { checked, value };
    }

    constructor(template: BaseTemplate<any>, options?: ValidationOptions) {
        super(template, options);
    }

}