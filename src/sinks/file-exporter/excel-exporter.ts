import exceljs, { Cell, Workbook } from "exceljs";
import { MappedRecord } from "../../transformer/types/transformer-dto.js";
import { TableTemplateOpts } from "../../template-mappers/types/table-template.type.js";
import { BaseTemplate } from "../../template-mappers/base-template.js";
import { FileExtension, FileOutputType, FileSink } from "../file-sinks.js";
export class ExcelExporter extends FileSink {
   
    protected templateStrct!: TableTemplateOpts;
    protected isSaveMetadata: boolean = false;
    protected workBook!: exceljs.Workbook;
    protected workSheet!: exceljs.Worksheet;

    constructor(
        typeOutput: FileOutputType,
        extension: FileExtension = FileExtension.EXCEL,
        template: BaseTemplate<TableTemplateOpts>
    ) {
        super(typeOutput, extension, template);
        this.templateStrct = template.getStructure();
        this.workBook = new Workbook();
        this.workSheet = this.workBook.addWorksheet();
    }

    async handle(chunk: MappedRecord[]): Promise<void> {
        for (let i = 0; i < chunk.length; i++) {
            const row = chunk[i];
            const rawRow = this.getRowValues(row);
            this.workSheet.addRow(rawRow);
        }
    }

    getRowValues(row: any) {
        const values: any[] = [];
        for (let i = 0; i < this.templateStrct.fields.length; i++) {
            const field = this.templateStrct.fields[i];
            const value = row[field.name]
            if(value === undefined) continue;
            values.push(value);
        }

        return values;
    }
}