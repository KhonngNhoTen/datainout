import { CellDataHelper, ExcelReaderHelper, RowDataHelper, SheetDataHelper } from "../../../helpers/excel-reader-helper.js";
import { TypeParser } from "../../../helpers/parse-type.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";
import { IReader } from "../../../common/decorators/IReader.decorator.js";

export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ExcelReaderHelper = new ExcelReaderHelper();

  private cells: CellDataHelper[] = [];

  constructor() {
    super({ type: "excel", typeParser: new TypeParser() });
  }

  async load(arg: unknown): Promise<any> {
    this.excelReaderHelper = new ExcelReaderHelper({
      onSheet: async (sheet) => await this.readSheet(sheet),
      onRow: async (row) => await this.readRow(row),
      onCell: async (cell) => await this.readCell(cell),
      beginTables: this.templates.map((e) => e.beginTableAt),
      endTableAts: this.templates.map((e) => e.endTableAt),
      columnIndexes: this.templates.map((e) => e.keyTableAt),
    });

    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }

  private async readCell(cellRaw: CellDataHelper) {
    this.cells.push(cellRaw);
  }

  private async readRow(row: RowDataHelper) {
    if (this.cells.length === 0) return;
    const sheet = this.templates[this.sheetIndex];
    const cellDesc = this.groupCellDescs[row.section];
    const isTrigger = this.tableDataImportHelper.push(this.cells, sheet, cellDesc, this.chunkSize);

    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: row.section,
        sheetIndex: row.rowIndex,
        sheetName: sheet.sheetName,
      };
      const data = this.tableDataImportHelper.pop();
      await this.callHandlers(data, filter);
    }
    this.cells = [];
  }

  private async readSheet(sheet: SheetDataHelper) {
    const filter: FilterImportHandler = {
      section: "footer",
      sheetIndex: this.sheetIndex,
      sheetName: sheet.name,
    };
    const data = this.tableDataImportHelper.pop();
    await this.callHandlers(data, filter);
  }
}
