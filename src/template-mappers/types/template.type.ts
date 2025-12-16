import { RawRecordType, SourceType } from "../../source/types/type.js";
import { ValidateResult } from "../../validators/types/type.js";

export type TableScope = "metadata"|"table";
export type TemplateField = {
        name: string;
        type: string;
        scope: TableScope
        required?: boolean;
        validate?: (data: any) => ValidateResult;
        setValue?: (data: any) => any;
}

export type Template = {
    fields: TemplateField[];
    metadata: TemplateField[];
    name?: string;
    type: RawRecordType;
    id: string;
    number: number;
    sourceType: SourceType
}