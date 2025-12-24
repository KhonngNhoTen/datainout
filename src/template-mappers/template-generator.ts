import { TABLE_KEY } from "../common/constant/constant.js";
import { EngineStream } from "../engines/engine.js";
import { PathHelper } from "../helpers/path.helper.js";
import { ExcelHelper } from "../helpers/table.helper.js";
import { ExcelReader } from "../readers/excel-readers.js";
import { ReaderResult } from "../readers/ireader.js";
import { TableRowRaw } from "../readers/type.js";
import { FileExtension } from "../sinks/export-sinks.js";
import { TableTemplateOpts, TemplateCellOpts } from "./types/table-template.type.js";
import fs from "fs";

type TemplateGeneratorContext = {
    visitedTableScope: "pending" | "visitting" | "visitted";
    source: "import" | "report";
    outputFilePath: string
}
export class TemplateGenerator {

    private static context: TemplateGeneratorContext = {
        visitedTableScope: "pending",
        source: "import",
        outputFilePath: ""
    }

    static async create(filePath: string, outputFilePath: string, source: "report" | "import"): Promise<void>
    static async create(buffer: Buffer, outputFilePath: string, source: "report" | "import"): Promise<void>
    static async create(arg: any, outputFilePath: string, source: "report" | "import" = "report") {

        TemplateGenerator.context = {
            visitedTableScope: "pending",
            source,
            outputFilePath
        }

        const template: TableTemplateOpts = {
            fields: [],
            metadata: [],
            id: 0 + "",
            type: "table", // Added missing property
            sourceType: "excel", // Added missing property
            number: 0,
            table: {
                startAt: -1,
                endAt: undefined
            }
        };
        const engineStream = new EngineStream({
            onChunk: (row: any) => TemplateGenerator.onRow(row, template),
            onEnd: () => TemplateGenerator.onEnd(template, outputFilePath),
        });


        const reader = new ExcelReader(arg, 1);
        reader.stream().pipe(engineStream);
        await reader.open();
        await reader.start();
        await engineStream.waitingDone();
    }

    private static async onRow(data: TableRowRaw, template: TableTemplateOpts, style: boolean = false) {
        const templateCells: TemplateCellOpts[] = [];

        data.cells.forEach(cell => {

            if (ExcelHelper.isVariableTable(cell.value) && TemplateGenerator.context.visitedTableScope === "pending")
                TemplateGenerator.context.visitedTableScope = "visitting";
            if (ExcelHelper.isVariableCell(cell.value) && TemplateGenerator.context.visitedTableScope === "visitting")
                TemplateGenerator.context.visitedTableScope = "visitted";

            let templateCell!: TemplateCellOpts;
            const isVariable = ExcelHelper.isVariableCell(cell) || ExcelHelper.isVariableTable(cell.value);
            if (isVariable || (!isVariable && style)) {
                const scope = TemplateGenerator.getVisittedTable(cell.value) === "visitted" ? "fields" : "metadata";
                const addressDetail = TemplateGenerator.splitAddress(cell.address);
                templateCell = {
                    address: cell.address,
                    addressDetail,
                    value: cell.value,
                    isVariable,
                    scope,
                    // required: cell.required,
                    // validate: cell.validate,
                    // setValue: cell.setValue,
                } as any;

                if (style)
                    templateCell = {
                        ...templateCell,
                        formula: cell.formula,
                        alignment: cell.alignment,
                        border: cell.border,
                        font: cell.font,
                        fill: cell.fill,
                        protection: cell.protection,
                        numFmt: cell.numFmt,
                    }
            }

            if (isVariable) {
                const parseCell = TemplateGenerator.parseCellVariable(cell.value);
                templateCell.type = parseCell.type;
                templateCell.name = parseCell.fieldName;
                delete templateCell.value;
            }

            templateCells.push(templateCell);
        });

        const fieldCells = templateCells.filter(cell => cell?.scope === "fields");
        const metadataCells = templateCells.filter(cell => cell?.scope === "metadata");
        if (fieldCells?.length > 0) template.fields.push(...fieldCells);
        if (metadataCells?.length > 0) template.metadata.push(...metadataCells);


        if (style) {
            template.alignment = data.alignment;
            template.border = data.border;
            template.font = data.font;
            template.fill = data.fill;
            template.height = data.height;
            template.protection = data.protection;
            template.numFmt = data.numFmt;
        }
    }

    private static async onEnd(template: TableTemplateOpts, outputFilePath: string, source: "report" | "import" = "report") {
        if (template.table.startAt === -1) template.table.startAt = 1;
        outputFilePath = PathHelper.getPath(source as any, outputFilePath, "templateDir");

        const context = `module.exports = ${JSON.stringify(template, null, 2)}`;
        fs.writeFileSync(outputFilePath, context);
    }

    private static getVisittedTable(cellValue: any) {
        const isVariableTable = ExcelHelper.isVariableTable(cellValue);
        const isVariable = ExcelHelper.isVariableCell(cellValue);

        if (TemplateGenerator.context.visitedTableScope === "pending" && isVariableTable)
            return "visitting";
        if (TemplateGenerator.context.visitedTableScope === "visitting" && !isVariable)
            return "visitted";
        return "pending";
    }

    private static parseCellVariable(cellValue: string) {
        cellValue = cellValue.replace("$$", "$");
        let fieldName = "";
        let type = "string";
        fieldName = cellValue.split("$")[1];
        if (fieldName.includes("->")) {
            const args = fieldName.split("->");
            fieldName = args[0];
            type = args[1].toLowerCase();
        }

        return { fieldName, type };
    }

    private static splitAddress(address: string) {
        const col = address.split(/\d+/)[0];
        const row = address.split(/[a-zA-Z]/)[1];
        return { col, row: Number(row) };
    }
}