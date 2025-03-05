import * as exceljs from "exceljs";
import { CellDescription, FilterImportHandler, SheetSection, TableData } from "../type.js";
import { SheetDesciption } from "../importers/reader/ImporterFileDescription.js";

type HandlerTableData = (data: ChunkData, filter: FilterImportHandler) => any;
export type ChunkData = { table?: exceljs.Cell[][]; header?: exceljs.Cell[]; footer?: exceljs.Cell[] };

type TableDataHandlerOptions = {
  beginTableAt: number;
  keyIndex?: number;
  chunkSize?: number;
  sheetIndex: number;
  sheetName: string;
  sheetDesciption: SheetDesciption;
};
export class TableDataHandler {
  private _beginTableAt: number;
  private _endTableAt: number = -1;
  private _events: {
    [key in keyof ChunkData]?: HandlerTableData;
  } = {};
  private _currentSection: SheetSection = "header";
  private _keyIndex: number = 0;
  private _chunkSize: number = 1;
  private _countRow = 0;
  private _filter: FilterImportHandler;
  private _data: ChunkData = { footer: [], header: [], table: [] };
  private _cellDescriptions: Record<keyof ChunkData, CellDescription[]> = { footer: [], header: [], table: [] };

  constructor(opts: TableDataHandlerOptions) {
    this._beginTableAt = opts.beginTableAt;
    this._keyIndex = opts.keyIndex ?? 1;
    this._chunkSize = opts.chunkSize ?? 1;

    this._filter = {
      section: "table",
      sheetIndex: opts.sheetIndex,
      sheetName: opts.sheetName,
    };

    this._cellDescriptions["header"] = opts.sheetDesciption.content.filter((e) => e.section === "header");
    this._cellDescriptions["table"] = opts.sheetDesciption.content.filter((e) => e.section === "table");
    this._cellDescriptions["footer"] = opts.sheetDesciption.content.filter((e) => e.section === "footer");
  }

  private getSection(row: exceljs.Row): SheetSection {
    let section: SheetSection = "table";
    if (row.number < this._beginTableAt) section = "header";
    else if (this._currentSection === "footer" || (this._currentSection === "table" && this.isEndTable(row))) section = "footer";

    if (section === "footer" && this._endTableAt === -1) {
      this._endTableAt = row.number - 1;
    }
    return section;
  }

  private setSection(section: SheetSection, row: exceljs.Row) {
    if (section !== this._currentSection || (section === "table" && this._countRow >= this._chunkSize)) {
      this.callEvents(this._currentSection);
    }

    this._currentSection = section;
  }

  addRow(row: exceljs.Row | null) {
    if (row === null) {
      this.callEvents(this._currentSection);
      return;
    }
    const section = this.getSection(row);
    this.setSection(section, row);

    if (this._beginTableAt === row.number) return;
    if (section === "table") this._countRow++;

    this.pushCells(row);
  }

  on(key: keyof TableData, handler: HandlerTableData) {
    this._events[key] = handler;
  }

  private callEvents(key: keyof TableData) {
    if (this._events[key]) {
      this._events[key](this._data, { ...this._filter, section: key });
    }

    this._data = {
      footer: [],
      header: [],
      table: [],
    };

    this._countRow = 0;
  }

  private isEndTable(row: exceljs.Row) {
    return !(row.values as [])[this._keyIndex + 1];
  }

  private pushCells(row: exceljs.Row) {
    const cells: exceljs.Cell[] = this.filterCells(row, this._currentSection);

    const gCells = cells.filter(
      (cell) => !((cell as any)._value.model.type === exceljs.ValueType.Merge) && cell.value !== undefined
    );
    if (gCells.length > 0) {
      if (this._currentSection === "table") this._data[this._currentSection]?.push(gCells);
      else this._data[this._currentSection]?.push(...gCells);
    }
  }

  filterCells(row: exceljs.Row, section: SheetSection): exceljs.Cell[] {
    const cells: exceljs.Cell[] = (row as any)._cells;
    let addresses: string[] = [];

    if (section === "header") {
      addresses = this._cellDescriptions.header.map((e) => e.address ?? "");
    } else if (section === "table")
      addresses = this._cellDescriptions.table.map((e) => {
        return `${e.address}${row.number}`;
      });
    else {
      addresses = this._cellDescriptions.footer.map((e) => {
        return `${e.address}${this._endTableAt + e.fullAddress.row}`;
      });
    }

    return cells.filter((cell) => addresses.includes(cell.address));
  }
}
