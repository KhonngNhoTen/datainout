import { EngineStream } from "../engines/engine.js";
import { PathHelper } from "../helpers/path.helper.js";
import { ExcelHelper } from "../helpers/table.helper.js";
import { ExcelReader } from "../readers/excel-readers.js";
import { TableRowRaw } from "../readers/type.js";
import fs from "fs";
import { TableTemplateOpts, TemplateCellOpts } from "../template-mappers/types/table-template.type.js";
import { Cell } from "exceljs";

type TemplateGeneratorContext = {
    visitedTableScope: "pending" | "visitting" | "visitted";
    source: "import" | "report";
    outputFilePath: string
}
export class ExcelTemplateGenerator {

    private static context: TemplateGeneratorContext = {
        visitedTableScope: "pending",
        source: "import",
        outputFilePath: ""
    }

    static async create(filePath: string, outputFilePath: string, source: "report" | "import"): Promise<void>
    static async create(buffer: Buffer, outputFilePath: string, source: "report" | "import"): Promise<void>
    static async create(arg: any, outputFilePath: string, source: "report" | "import" = "report") {

        ExcelTemplateGenerator.context = {
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
            onChunk: (row: any) => ExcelTemplateGenerator.onRow(row, template, source === "report"),
            onEnd: () => ExcelTemplateGenerator.onEnd(template),
        });

        const reader = new ExcelReader(arg, 0, 0);
        reader.stream().pipe(engineStream);
        await reader.open();
        await reader.start();
        await engineStream.waitingDone();
    }

    private static async onRow(data: TableRowRaw, template: TableTemplateOpts, style: boolean = false) {
        const templateCells: TemplateCellOpts[] = [];

        data.cells.forEach(cell => {

            if (ExcelHelper.isVariableTable(cell.value) && ExcelTemplateGenerator.context.visitedTableScope === "pending")
                ExcelTemplateGenerator.context.visitedTableScope = "visitting";
            if (ExcelHelper.isVariableCell(cell.value) && ExcelTemplateGenerator.context.visitedTableScope === "visitting")
                ExcelTemplateGenerator.context.visitedTableScope = "visitted";

            let templateCell!: TemplateCellOpts;
            const isVariable = ExcelHelper.isVariableCell(cell) || ExcelHelper.isVariableTable(cell.value);
            if (isVariable || (!isVariable && style)) {
                const scope = ExcelTemplateGenerator.getVisittedTable(cell.value) === "visitted" ? "fields" : "metadata";
                const addressDetail = ExcelTemplateGenerator.splitAddress(cell.address);
                (addressDetail as any).columnIndex = cell.fullAddress.col;
                templateCell = {
                    address: cell.address,
                    addressDetail,
                    value: cell.value,
                    isVariable,
                    scope,
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
                const parseCell = ExcelTemplateGenerator.parseCellVariable(cell.value);
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

        if(template.table.startAt === -1 && ExcelTemplateGenerator.context.visitedTableScope === "visitting")
            template.table.startAt = data.number;

        // if (style) {
        //     template.alignment = data.alignment;
        //     template.border = data.border;
        //     template.font = data.font;
        //     template.fill = data.fill;
        //     template.height = data.height;
        //     template.protection = data.protection;
        //     template.numFmt = data.numFmt;
        // }
    }

    private static async onEnd(template: TableTemplateOpts) {
        const source =  ExcelTemplateGenerator.context.source;
        
        let outputFilePath =  ExcelTemplateGenerator.context.outputFilePath;
        if (template.table.startAt === -1) template.table.startAt = 1;
        outputFilePath = PathHelper.getPath(source as any, outputFilePath, "templateDir");

        const context = `module.exports = ${JSON.stringify(template, null, 2)}`;
        fs.writeFileSync(outputFilePath, context);
    }

    private static getVisittedTable(cellValue: any) {
        const isVariableTable = ExcelHelper.isVariableTable(cellValue);
        const isVariable = ExcelHelper.isVariableCell(cellValue);

        if (ExcelTemplateGenerator.context.visitedTableScope === "pending" && isVariableTable)
            return "visitting";
        if (ExcelTemplateGenerator.context.visitedTableScope === "visitting" && !isVariable)
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
        const column = address.split(/\d+/)[0];
        const row = address.split(/[a-zA-Z]/)[1];
        return { column, row: Number(row) };
    }
}