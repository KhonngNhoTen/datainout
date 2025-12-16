import { IReader } from "./ireader.js";
import exceljs from "exceljs";
import { ReaderOpenOpts, TableRowRaw } from "./type.js";

export class ExcelReader implements IReader {
   
    private workBook!: exceljs.Workbook;
    private metadata: any = {};

    async open(options?: ReaderOpenOpts): Promise<void> {
        this.workBook = new exceljs.Workbook();
        if (options?.filePath) await this.workBook.xlsx.readFile(options?.filePath);
        if (options?.buffer) await this.workBook.xlsx.load(options.buffer as unknown as exceljs.Buffer);
    }

    getIterator(): AsyncIterableIterator<TableRowRaw> {
        const workSheet = this.workBook.getWorksheet();

        async function* iterator() {
            if (workSheet === undefined) throw new Error("Not found worksheet in file");
            for (let i = 1; i < workSheet.rowCount; i++) {
                const row = workSheet.getRow(i);
                const cells: exceljs.Cell[] = [];
                row.eachCell(c => cells.push(c));                
                yield {
                    cells,
                    actualCellCount: row.actualCellCount,
                    cellCount: row.cellCount,
                    height: row.height,
                    id: row.number,
                    number: row.number,
                    model: row.model,
                    type: 'table',
                    outlineLevel: row.outlineLevel
                } as unknown as TableRowRaw;
            }
        }
        return iterator();
    }

    cancel(): void {
    }

    async close() {
    }
}