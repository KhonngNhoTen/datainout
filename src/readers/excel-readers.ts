import { AddCallback, IReader, ReaderResult } from "./ireader.js";
import * as exceljs from "exceljs";
import { _TableRowRaw, TableRowRaw } from "./type.js";
import { ExcelHelper } from "../helpers/table.helper.js";

export class ExcelReader extends IReader {

    private filePath!: string;
    private buffer!: Buffer;
    private workBook!: exceljs.Workbook;
    protected startAt: number;
    protected endAt?: number;


    constructor(filePath: string, startAt: number, endAt?: number)
    constructor(buffer: Buffer, startAt: number, endAt?: number)
    constructor(arg1: any, startAt: number, endAt?: number) {
        super();
        if (typeof arg1 === "string") this.filePath = arg1;
        else if (arg1 instanceof Buffer) this.buffer = arg1;
        this.startAt = startAt;
        this.endAt = endAt;
    }


    async open(): Promise<void> {
        this.workBook = new exceljs.Workbook();
        if (this?.filePath) { 
            await this.workBook.xlsx.readFile(this?.filePath);
        }
        else if (this?.buffer) {
            await this.workBook.xlsx
                .load(this.buffer as unknown as exceljs.Buffer)
        }
    }

    async get(add: AddCallback) {
        const workSheet = this.workBook.getWorksheet();

        if (workSheet === undefined) throw new Error("Not found worksheet in file");
        for (let i = 1; i <= workSheet.rowCount; i++) {
            const _row = workSheet.getRow(i) ;
            const row = _row as unknown as _TableRowRaw;
            row.type = "table";
            row.scope = ExcelHelper.getScope(i, this.startAt);
            row.cells = (row as any)._cells;
            add(row);
        }

        add(null);
    }

}