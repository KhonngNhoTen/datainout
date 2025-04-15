import { CellDataHelper, ExcelReaderHelper, RowDataHelper, SheetDataHelper } from "../../../helpers/excel-reader-helper.js";
import { TypeParser } from "../../../helpers/parse-type.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";
import { IReader } from "../../../common/decorators/IReader.decorator.js";

@IReader()
export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ExcelReaderHelper;

  private cells: CellDataHelper[] = [];

  constructor() {
    super({ type: "excel", typeParser: new TypeParser() });
    this.excelReaderHelper = new ExcelReaderHelper({
      onSheet: async (sheet) => await this.readSheet(sheet),
      onRow: async (row) => await this.readRow(row),
      onCell: async (cell) => await this.readCell(cell),
    });
  }

  async load(arg: unknown): Promise<any> {
    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }

  private async readCell(cellRaw: CellDataHelper) {
    this.cells.push(cellRaw);
  }

  private async readRow(row: RowDataHelper) {
    const sheet = this.templates[this.sheetIndex];
    const cellDesc = this.groupCellDescs[row.section];
    const isTrigger = this.tableDataImportHelper.push(this.cells, sheet, cellDesc, this.chunkSize);

    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: row.section,
        sheetIndex: row.rowIndex,
        sheetName: sheet.sheetName,
      };
      const data = this.tableDataImportHelper.get();
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
    const data = this.tableDataImportHelper.get();
    await this.callHandlers(data, filter);
  }
}
