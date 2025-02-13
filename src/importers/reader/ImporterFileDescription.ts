import { CellDescription, TemplateExcelImportOptions, SheetDesciptionOptions, SheetSection } from "../type.js";

export class SheetDesciption {
  beginTableAt: number;
  endTableAt?: number;
  content: Array<CellDescription>;
  index: number;
  keyIndex: number = 1;
  name?: string;
  constructor(opts: SheetDesciptionOptions, index: number) {
    this.beginTableAt = opts.beginTableAt ?? 1;
    this.endTableAt = opts.endTableAt;
    this.content = opts.content;
    this.index = index;
    this.name = opts.name;
    this.keyIndex = opts.keyIndex ?? 1;
  }

  findCellByAddress(address: string, rowIndex: number) {
    let section: SheetSection = "table";
    if (rowIndex < this.beginTableAt) section = "header";
    else if (this.endTableAt && rowIndex > this.endTableAt) section = "footer";

    const foundIndex = this.content.findIndex(
      (e) =>
        e.address && ((section !== "table" && address === e.address) || (section === "table" && address.indexOf(e.address) === 0))
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
