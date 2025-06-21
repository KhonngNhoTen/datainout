import * as exceljs from "exceljs";
import { SheetSection, TableData } from "../common/types/common-type.js";
import { RowDataHelper } from "../common/types/excel-reader-helper.type.js";
import { CellImportOptions, SheetImportOptions } from "../common/types/import-template.type.js";
import { TypeParser } from "./parse-type.js";
import { DEFAULT_BEGIN_TABLE, DEFAULT_END_TABLE, ReaderExceljsHelper } from "./excel.helper.js";
import { validateCellImport } from "./validate-cell-importer.js";
import { ConvertorRows2TableDataOpts, GroupValueRow } from "../common/types/convert-row-to-table-data.type.js";

export class ConvertorRows2TableData {
  private chunkSize: number;
  private section?: SheetSection;
  private typeParser: TypeParser;
  private container: TableData = { header: {}, footer: {}, table: [] };
  private beginTable: number = DEFAULT_BEGIN_TABLE;
  private endTable: number = DEFAULT_END_TABLE;

  constructor(opts?: ConvertorRows2TableDataOpts) {
    this.chunkSize = opts?.chunkSize ?? 10;
    this.typeParser = opts?.typeParser ?? new TypeParser();
  }

  push(row: RowDataHelper | null, template: SheetImportOptions): { isTrigger: boolean; triggerSection: SheetSection };
  push(row: exceljs.Row | null, template: SheetImportOptions): { isTrigger: boolean; triggerSection: SheetSection };
  push(arg: any, template: SheetImportOptions) {
    if (arg === null) return { isTrigger: true, triggerSection: this.section ?? "header" };
    if (ReaderExceljsHelper.isNullableRow(arg?.detail ?? arg)) return { isTrigger: false, triggerSection: this.section ?? "header" };

    const { endTable, row, section } = this.getRowInformation(arg, template);
    const isTrigger = this.isTrigger(section);
    const triggerSection = this.triggerSection(section);

    // Map cell velue with cell description
    let groupValues: GroupValueRow = {};
    if (!row) return { isTrigger: false, triggerSection: this.section ?? "header" };

    for (let i = 0; i < row.cellCount; i++) {
      const cell = row.getCell(i + 1);
      if (cell?.value === undefined) continue;
      const index = template.cells.findIndex((e) => this.compareByAddress(e, section, cell, row.number, endTable));
      if (index === -1) continue;
      const cellImport = template.cells[index];
      groupValues[cellImport.keyName] = cell.value;
    }

    if (isTrigger || section === "table") {
      groupValues = this.formatValue(
        template.cells.filter((e: CellImportOptions) => e.section === section),
        groupValues,
        this.typeParser,
        row.number
      );
    }

    if (section === "table" && Object.keys(groupValues).length > 0) this.container.table?.push(groupValues);
    else if (Object.keys(groupValues).length > 0)
      this.container[section] = {
        ...this.container[section],
        ...(groupValues as any),
      };

    this.section = section;
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
      const chunkSize = this.container?.table?.length;
      if (this.chunkSize - 1 === chunkSize && section === this.section && section === "table") return true;
    }
    return false;
  }

  private triggerSection(section: SheetSection) {
    let triggerSection: SheetSection = section;
    if (!this.isTrigger(section)) return section;

    if (this.section && this.section !== section) {
      if (section === "header") triggerSection = "header";
      if (section === "table") triggerSection = "header";
      if (section === "footer") triggerSection = "table";
    }

    if (this.section && section === this.section && this.section === "table") triggerSection = "table";
    return triggerSection;
  }

  private getRowInformation(arg: any, template: SheetImportOptions) {
    let row: exceljs.Row,
      section: SheetSection = "header";
    if (arg.detail) {
      row = arg.detail;
      section = arg.section;
    } else row = arg;

    let { beginTable, endTable } = this;
    const drafBeginTable = ReaderExceljsHelper.beginTableAt(row, template, false);
    const drafEndTable = ReaderExceljsHelper.endTableAt(row, template, false);

    if (beginTable === DEFAULT_BEGIN_TABLE && drafBeginTable) beginTable = drafBeginTable;
    if (endTable === DEFAULT_END_TABLE && drafEndTable) endTable = drafEndTable;

    if (!arg.detail) section = ReaderExceljsHelper.getSection(row, beginTable, endTable);
    this.endTable = endTable;
    this.beginTable = beginTable;
    return { row, section, beginTable, endTable };
  }

  private formatValue(formattedCellImport: CellImportOptions[], groupValues: GroupValueRow, typeParser: TypeParser, rowIndex: number) {
    const row = { ...groupValues };
    formattedCellImport.forEach((cell) => {
      let value = groupValues[cell.keyName];
      if (cell.setValue) value = cell.setValue(value, row);
      if (cell.type && cell.type !== "virtual") {
        value = (typeParser as any)[cell.type](value);
        return;
      }
      const { col } = ReaderExceljsHelper.splitAddress(cell.address ?? "");
      validateCellImport(value, cell, { row: rowIndex, col });
      groupValues[cell.keyName] = value;
    });

    return groupValues;
  }

  private compareByAddress(cellDes: CellImportOptions, section: SheetSection, cell: exceljs.Cell, rowNumber: number, endTableAt: number) {
    if (cellDes.section === section) {
      if (cellDes.section === "header") return cellDes.address === cell.address;
      else if (cellDes.section === "table") return cell.address === `${cellDes.address}${rowNumber}`;
      else {
        return rowNumber === endTableAt + (cellDes?.fullAddress?.row ?? 0);
      }
    }
    return false;
  }
}
