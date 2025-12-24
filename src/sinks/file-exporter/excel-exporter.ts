import exceljs, { Cell, Workbook } from "exceljs";
import { MappedRecord } from "../../transformer/types/transformer-dto.js";
import { TableTemplateOpts, TemplateCellOpts } from "../../template-mappers/types/table-template.type.js";
import { BaseTemplate } from "../../template-mappers/base-template.js";
import { FileExtension, FileOutputType, ExportSink } from "../export-sinks.js";
export class ExcelExporter extends ExportSink<"buffer" | "file"> {
   
    protected templateStrct!: TableTemplateOpts;
    protected isSaveMetadata: boolean = false;
    protected workBook!: exceljs.Workbook;
    protected workSheet!: exceljs.Worksheet;
    protected isWroteHeader: boolean = false;

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
        if(!this.isWroteHeader) this.writeHeader(chunk[0].metadata);
        for (let i = 0; i < chunk.length; i++) {
            const row = chunk[i];
            const rawRow = this.getRowValues(row.fields);
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

    private writeHeader(metadata: any) {
        const headers: Record<string, TemplateCellOpts[]> = {};
        let lastRow = 1;
        this.templateStrct.metadata.forEach(m => {
            if(headers[m.addressDetail.row]) headers[m.addressDetail.row].push(m);
            else headers[m.addressDetail.row] = [m];
            if(m.addressDetail.row > lastRow) lastRow = m.addressDetail.row;
        });   


        for (let i = 1; i <= lastRow; i++) {
            const row = this.workSheet.addRow([]);
            headers[i].forEach(e => console.log(e.addressDetail, e.value));
            if(headers[i]) {
                headers[i].forEach((h: TemplateCellOpts) => {
                    const cell = row.getCell(h.addressDetail.columnIndex);
                    cell.value = h.isVariable ? metadata[h.name] : h.value;
                    cell.style = h;
                })
            }
        }

        this.isWroteHeader = true;
    }

    async export() {
        if(this.typeOutput==="buffer") return await this.workBook.xlsx.writeBuffer();
        return await this.workBook.xlsx.writeFile(this.filePath) as any;    
    }
}