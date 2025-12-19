import { AddCallback, IReader, ReaderResult } from "./ireader.js";
import exceljs from "exceljs";
import { TableRowRaw } from "./type.js";

export class ExcelReader extends IReader {

    private filePath!: string;
    private buffer!: Buffer;

    constructor(filePath: string)
    constructor(buffer: Buffer)
    constructor(arg1: any) {
        super();
        if (typeof arg1 === "string") this.filePath = arg1;
        else if (arg1 instanceof Buffer) this.buffer = arg1;
    }

    private workBook!: exceljs.Workbook;

    async open(): Promise<void> {
        this.workBook = new exceljs.Workbook();
        if (this?.filePath) await this.workBook.xlsx.readFile(this?.filePath);
        if (this?.buffer) await this.workBook.xlsx.load(this.buffer as unknown as exceljs.Buffer);
    }

    async get(add: AddCallback) {
        const workSheet = this.workBook.getWorksheet();

        if (workSheet === undefined) throw new Error("Not found worksheet in file");
        for (let i = 1; i < workSheet.rowCount; i++) {
            const row = workSheet.getRow(i);
            const cells: exceljs.Cell[] = [];
            row.eachCell(c => cells.push(c));
            add({
                cells,
                actualCellCount: row.actualCellCount,
                cellCount: row.cellCount,
                height: row.height,
                id: row.number,
                number: row.number,
                model: row.model,
                type: 'table',
                outlineLevel: row.outlineLevel
            } as unknown as TableRowRaw);
        }

        add(null);
    }

}