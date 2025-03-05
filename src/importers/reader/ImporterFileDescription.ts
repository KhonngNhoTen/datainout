import { CellDescription, TemplateExcelImportOptions, SheetDesciptionOptions, SheetSection } from "../type.js";

export class SheetDesciption {
  startTable: number;
  endTable?: number;
  content: Array<CellDescription>;
  index: number;
  keyIndex: number = 1;
  name?: string;
  constructor(opts: SheetDesciptionOptions, index: number) {
    this.startTable = opts.startTable ?? 1;
    this.endTable = opts.endTable;
    this.content = opts.content;
    this.index = index;
    this.name = opts.name;
    this.keyIndex = opts.keyIndex ?? 1;
  }

  findCellByAddress(address: string, rowIndex: number) {
    let section: SheetSection = "table";
    if (rowIndex < this.startTable) section = "header";
    else if (this.endTable && rowIndex > this.endTable) section = "footer";

    const foundIndex = this.content.findIndex(
      (e) =>
        e.address &&
        ((section !== "table" && address === e.address) || (section === "table" && address.indexOf(e.address) === 0)),
    );
    if (foundIndex === -1) return null;
    return this.content[foundIndex];
  }
}

export class ImportFileDesciption {
  sheets: SheetDesciption[];
  constructor(opts: TemplateExcelImportOptions) {
    this.sheets = opts.sheets.map((val, index) => new SheetDesciption(val, index));
  }
}
