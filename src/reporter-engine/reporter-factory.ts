import { PathHelper } from "../helpers/path.helper.js";
import { FileExtension, FileOutputType, FileSink } from "../sinks/file-sinks.js";
import { BaseSource } from "../source/base-source.js";
import { BaseTemplate } from "../template-mappers/base-template.js";
import { Template } from "../template-mappers/types/template.type.js";
import { Reporter } from "./reporter.js";
import { Validators } from "../validators/validators.js";
import { ExcelExporter } from "../sinks/file-exporter/excel-exporter.js";
import fs from "fs";

type ReportFactoryOptions = {
    extension: FileExtension,
    typeOutput: FileOutputType,
    source: BaseSource,
    template: string | BaseTemplate<Template>,
    errorStrategy?: "fail-fast" | "skip" | "collect" | "redirect-file",
}

export class ReportFactory {
    static create(pattern: ReportFactoryOptions): Reporter<any>
    {
        const template: BaseTemplate<any> =
            typeof pattern.template === "string" ?
                this.createTemplate(pattern.template) :
                pattern.template;

        const source = pattern.source;
        const sink = this.createSink(pattern.extension, pattern.typeOutput, template);

        return new Reporter(source, template, new Validators(template, { errorStrategy: pattern.errorStrategy }), sink) as any;
    }

    protected static createTemplate(templatePath: string) {
        const tempPath = PathHelper.getPath("report", templatePath, "templateDir");
        let template!: BaseTemplate<any>;

        if (!fs.existsSync(tempPath)) template = new BaseTemplate();
        else {
            const templOpts: Template = require(tempPath);
            template = new BaseTemplate(templOpts);
        }
        return template;
    }

    protected static createSink(
        ext: FileExtension,
        typeOutput: FileOutputType,
        template: BaseTemplate<any>,
    ) {
        return new ExcelExporter(typeOutput, ext, template);
    }
}