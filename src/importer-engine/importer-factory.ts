import { EngineFactory, PatternEngine, SourceOptions } from "../engines/engine-factory.js";
import { BaseSource, isBaseSource } from "../source/base-source.js";
import { SourceAdapter } from "../source/source.js";
import { SourceType } from "../source/types/type.js";
import { TableTemplateOpts } from "../template-mappers/types/table-template.type.js";
import { Template } from "../template-mappers/types/template.type.js";
import { Importer } from "./importer.js";


type PatternImporter = Omit<PatternEngine, "source"> & {
    source: string | BaseSource | SourceOptions
};

export class ImporterFactory extends EngineFactory {
    static create(type: "excel", pattern: PatternImporter): Importer<TableTemplateOpts>
    static create(type: "csv", pattern: PatternImporter): Importer<TableTemplateOpts>
    static create(type: "log", pattern: PatternImporter): Importer<Template>
    static create(type: SourceType, pattern: PatternImporter): Importer<Template> {
        const template = ImporterFactory.createTemplate(type, pattern.templatePath);

        let source!: BaseSource;
        
        if (isBaseSource(pattern.source)){ 
            source = pattern.source;
            source.init(template.getStructure(), {});
        }
        else {
            const opts = {
                batchSize: typeof pattern.source === "string" ? undefined : pattern?.source?.batchSize,
                file: typeof pattern.source === "string" ? pattern.source : pattern.source?.file,
                concurrency: typeof pattern.source === "string" ? undefined : pattern.source?.concurrency,
                readerConfig: typeof pattern.source === "string" ? undefined : pattern.source?.readerConfig,
            }

            source = this.createSource(type, opts, template);
        }

        const validator = pattern.validator ?? ImporterFactory.createValidator(template);
        const transformer = pattern.transformer ?? ImporterFactory.createTransformer(template.getStructure().type, template);
        return new Importer(source, template, transformer, validator, pattern.sink, pattern.options)
    }
}