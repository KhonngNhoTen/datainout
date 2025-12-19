import { Readable } from "stream";
import exceljs from "exceljs";
import { AddCallback, IReader, ReaderResult } from "./ireader.js";
import EventEmitter from "events";
import { TableRowRaw } from "./type.js";

export class ExcelStreamReader extends IReader {

    protected streamable!: Readable;
    protected workBookReader!: exceljs.stream.xlsx.WorkbookReader;

    constructor(streamable: Readable) {
        super();
        this.workBookReader = new exceljs.stream.xlsx.WorkbookReader(streamable, {
        });
    }

    async get(add: AddCallback) {
        const workBookReaderEmitter = this.workBookReader as unknown as EventEmitter;
        workBookReaderEmitter.on("workSheet", async (worksheet: exceljs.Worksheet) => {
            const workSheetEmitter = worksheet as unknown as EventEmitter;
            workSheetEmitter.on("row", (row: exceljs.Row) => {
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
            });

            workSheetEmitter.on("finished", () => { })
        });

        workBookReaderEmitter.on("finished", () => {
            add(null);
        })

    }
}