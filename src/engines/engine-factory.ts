import { PathHelper } from "../helpers/path.helper.js";
import { ExcelReaderConfig } from "../readers/type.js";
import { Sink } from "../sinks/base-sinks.js";
import { SourceAdapter } from "../source/source.js";
import { RawRecordType, SourceType } from "../source/types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { Template } from "../template-mappers/types/template.type.js";
import { BaseTransformer } from "../transformer/base-transformer.js";
import { TableTransformer } from "../transformer/table-transformer.js";
import { BaseValidator } from "../validators/base-validator.js";
import { Validators } from "../validators/validators.js";
import { Engine } from "./engine.js";
import fs from "fs";

export type SourceOptions = {
    file: Buffer | string;
    batchSize?: number;
    concurrency?: number;
    readerConfig?: ExcelReaderConfig
}

export type PatternEngine = Record<string, any> & {
    source: SourceOptions | string,
    sink: Sink,
    templatePath: string,
    validator?: BaseValidator,
    transformer?: BaseTransformer<any>,
    options?: any
}

export class EngineFactory {
    static create(type: string, pattern: PatternEngine): Engine<any> {
        throw new Error("Method not implemented.");
    }

    protected static createTemplate(type: SourceType, templatePath: string) {
        const tempPath = PathHelper.getPath("import", templatePath);
        let template!: BaseTemplate<any>;

        if (!fs.existsSync(tempPath)) template = new BaseTemplate();
        else {
            const templOpts: Template = require(tempPath).template;
            if (templOpts.type === "table") template = new BaseTemplate(templOpts);
        }
        return template;
    }

    protected static createSource(type: SourceType, options: SourceOptions, template: BaseTemplate<any>) {
        let source!: SourceAdapter;
        if (type === "excel")
            source = new SourceAdapter(template.getStructure(), options);
        return source;
    }

    protected static createValidator(template: BaseTemplate<any>) {
        return new Validators(template)
    }

    protected static createTransformer(recordType: RawRecordType, template: BaseTemplate<any>) {
        if (recordType === "table") return new TableTransformer(template);
        return new TableTransformer(template);
    }
}