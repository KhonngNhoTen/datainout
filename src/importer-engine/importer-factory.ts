import { PathHelper } from "../helpers/path.helper.js";
import { ExcelReader } from "../readers/excel-readers.js";
import { IReader } from "../readers/ireader.js";
import { ExcelReaderConfig } from "../readers/type.js";
import { Sink } from "../sinks/base-sinks.js";
import { BaseSource, isBaseSource } from "../source/base-source.js";
import { Source } from "../source/source.js";
import { RawRecordType, SourceType } from "../source/types/type.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { Template } from "../template-mappers/types/template.type.js";
import { BaseTransformer } from "../transformer/base-transformer.js";
import { EventTransformer } from "../transformer/event-transformer.js";
import { TableTransformer } from "../transformer/table-transformer.js";
import { Validators } from "../validators/validators.js";
import { Importer } from "./importer.js";
import fs from "fs";


type ImporterFactoryOptions = {
    sink: Sink,
    source: BaseSource | string | Buffer,
    template: string | BaseTemplate<Template>,
    type?: SourceType,
    sourceOptions?: {
        batchSize?: number;
        concurrency?: number;
        readerConfig?: ExcelReaderConfig;
    },
    errorStrategy?: "fail-fast" | "skip" | "collect" | "redirect-file",
    validator?: Validators,
    transformer?: BaseTransformer<any>,

};

export class ImporterFactory {
    static create(pattern: Omit<ImporterFactoryOptions, "type"> & { type: "log" }): Importer<Template>
    static create(pattern: Omit<ImporterFactoryOptions, "type"> & { type: Exclude<SourceType, "log"> }): Importer<TableTemplateOpts>
    static create(pattern: Omit<ImporterFactoryOptions, "type">): Importer<TableTemplateOpts>
    static create(pattern: ImporterFactoryOptions): Importer<Template> {
        const type: SourceType = pattern.type ?? "excel";
        const template: BaseTemplate<any> =
            typeof pattern.template === "string" ?
                this.createTemplate(type, pattern.template) :
                pattern.template;

        const source = isBaseSource(pattern.source) ? pattern.source : this.createSource(type, pattern.source, template.getStructure());
        // source.init(template.getStructure());

        const validator = pattern.validator ?? new Validators(template, { errorStrategy: pattern.errorStrategy });
        const transformer = pattern.transformer ?? this.createTransformer(template.getStructure().type, template);

        return new Importer(source, template, transformer, validator, pattern.sink);
    }

    protected static createTemplate(type: SourceType, templatePath: string) {
        const tempPath = PathHelper.getPath("import", templatePath, "templateDir");
        let template!: BaseTemplate<any>;

        if (!fs.existsSync(tempPath)) template = new BaseTemplate();
        else {
            const templOpts: Template = require(tempPath);
            template = new BaseTemplate(templOpts);
        }
        return template;
    }

    protected static createSource(type: SourceType, inputSource: string | Buffer, template: Template) {
        let source!: Source;
        let reader!: IReader;
        if (type === "excel") {
            const tableTemplate = template as TableTemplateOpts;
            reader = new ExcelReader(inputSource as string, tableTemplate.table.startAt, tableTemplate.table.endAt);
        }

        source = new Source(template, reader);
        return source;
    }

    protected static createTransformer(recordType: RawRecordType, template: BaseTemplate<any>) {
        if (recordType === "table") return new TableTransformer(template);
        if (recordType === "event") return new EventTransformer(template);
        return new TableTransformer(template);
    }
}