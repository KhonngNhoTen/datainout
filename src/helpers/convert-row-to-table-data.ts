import * as exceljs from "exceljs";
import { SheetSection, TableData } from "../common/types/common-type.js";
import { RowDataHelper } from "../common/types/excel-reader-helper.type.js";
import { CellImportOptions, SheetImportOptions } from "../common/types/import-template.type.js";
import { TypeParser } from "./parse-type.js";
import { DEFAULT_BEGIN_TABLE, DEFAULT_END_TABLE, ReaderExceljsHelper } from "./excel.helper.js";

type ConvertorRows2TableDataOpts = {
  chunkSize?: number;
  typeParser?: TypeParser;
};
export class ConvertorRows2TableData {
  private chunkSize: number;
  private section?: SheetSection;
  private typeParser: TypeParser;
  private container: TableData = { header: {}, footer: {}, table: [] };

  constructor(opts?: ConvertorRows2TableDataOpts) {
    this.chunkSize = opts?.chunkSize ?? 10;
    this.typeParser = opts?.typeParser ?? new TypeParser();
  }

  push(row: RowDataHelper | null, template: SheetImportOptions): { isTrigger: boolean; triggerSection: SheetSection };
  push(row: exceljs.Row | null, template: SheetImportOptions): { isTrigger: boolean; triggerSection: SheetSection };
  push(arg: any, template: SheetImportOptions) {
    if (arg === null) return { isTrigger: true, triggerSection: this.section ?? "header" };

    const { endTable, row, section } = this.getRowInformation(arg, template);
    const isTrigger = this.isTrigger(section);
    const triggerSection = this.triggerSection(section);

    // Map cell velue with cell description
    let value: any = {};

    for (let i = 0; i < row.cellCount; i++) {
      const cell = row.getCell(i + 1);
      const index = template.cells.findIndex((e) => this.compareByAddress(e, section, cell, row.number, endTable));
      if (index === -1) continue;
      const cellImport = template.cells[index];
      value[cellImport.keyName] = cell.value;
    }

    if (isTrigger)
      value = this.formatValue(
        template.cells.filter((e: CellImportOptions) => e.section === section),
        value,
        this.typeParser
      );

    if (section === "table" && Object.keys(value).length > 0) this.container.table?.push(value);
    else if (Object.keys(value).length > 0) this.container[section] = { ...this.container[section], ...value };

    return { isTrigger, triggerSection };
  }

  pop(section: SheetSection) {
    const result = this.container[section];
    if (section === "table") this.container.table = [];
    else this.container[section] = {};
    return result;
  }

  private isTrigger(section: SheetSection) {
    if (this.section) {
      if (section !== this.section) return true;
      if (this.chunkSize === 0 && section === this.section && section === "table") return true;
    }
    return false;
  }

  private triggerSection(section: SheetSection) {
    let triggerSection: SheetSection = "header";

    if (this.section && this.section !== section) {
      if (this.section === "header") triggerSection = "header";
      if (this.section === "table") triggerSection = "header";
      if (this.section === "footer") triggerSection = "table";
    }

    if (this.section && section === this.section && this.section === "table") triggerSection = "table";

    return triggerSection;
  }

  private getRowInformation(arg: any, template: SheetImportOptions) {
    let row: exceljs.Row,
      section: SheetSection = "header";
    if (!arg.cells) {
      row = arg.detail;
      section = arg.section;
    } else row = arg;

    const beginTable = ReaderExceljsHelper.beginTableAt(row, template, false) ?? DEFAULT_BEGIN_TABLE;
    const endTable = ReaderExceljsHelper.endTableAt(row, template, false) ?? DEFAULT_END_TABLE;
    if (arg.cells) section = ReaderExceljsHelper.getSection(row, beginTable, endTable);
    return { row, section, beginTable, endTable };
  }

  private formatValue(formattedCellImport: CellImportOptions[], groupValues: any, typeParser: TypeParser) {
    const row = { ...groupValues };
    formattedCellImport.forEach((cell) => {
      let value = groupValues[cell.keyName];
      if (cell.setValue) value = cell.setValue(value, row);
      if (cell.type && cell.type !== "virtual") value = (typeParser as any)[cell.type](value);
      if (cell.validate && cell.validate(value)) throw new Error("Validated fail");

      groupValues[cell.keyName] = value;
    });

    return groupValues;
  }

  private compareByAddress(cellDes: CellImportOptions, section: SheetSection, cell: exceljs.Cell, rowNumber: number, endTableAt: number) {
    if (cellDes.section === section) {
      if (cellDes.section === "header") return cell.address === cell.address;
      else if (cellDes.section === "table") return cell.address === `${cell.address}${rowNumber}`;
      else {
        return rowNumber === endTableAt + (cell?.fullAddress?.row ?? 0);
      }
    }
    return false;
  }
}
