import { Readable } from "stream";
import exceljs from "exceljs";
import { AddCallback, IReader, ReaderResult } from "./ireader.js";
import EventEmitter from "events";
import { _TableRowRaw, TableRowRaw } from "./type.js";
import { ExcelHelper } from "../helpers/table.helper.js";

export class ExcelStreamReader extends IReader {

    protected streamable!: Readable;
    protected workBookReader!: exceljs.stream.xlsx.WorkbookReader;
    protected startAt: number;
    protected endAt?: number;

    constructor(streamable: Readable, startAt: number, endAt?: number) {
        super();
        this.workBookReader = new exceljs.stream.xlsx.WorkbookReader(streamable, {
        });
        this.startAt = startAt;
        this.endAt = endAt;
    }

    async get(add: AddCallback) {
        const workBookReaderEmitter = this.workBookReader as unknown as EventEmitter;
        workBookReaderEmitter.on("workSheet", async (worksheet: exceljs.Worksheet) => {
            const workSheetEmitter = worksheet as unknown as EventEmitter;
            workSheetEmitter.on("row", (row: _TableRowRaw) => {
                row.type = "table";
                row.scope = ExcelHelper.getScope(row.number, this.startAt);
                add(row);
            });

            workSheetEmitter.on("finished", () => { })
        });

        workBookReaderEmitter.on("finished", () => {
            add(null);
        })

    }
}