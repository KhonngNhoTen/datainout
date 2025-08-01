import * as exceljs from "exceljs";
import { ExcelTemplateManager } from "../common/core/Template.js";
import { ValidateImportError } from "../common/error/ValidateError.js";
import { EventType, SheetSection } from "../common/types/common-type.js";
import { CellImportOptions } from "../common/types/import-template.type.js";
import { TableData } from "../common/types/template.type.js";
import { DEFAULT_BEGIN_TABLE, DEFAULT_END_TABLE, ReaderExceljsHelper } from "./excel.helper.js";
import { TypeParser } from "./parse-type.js";
import { validateCellImport } from "./validate-cell-importer.js";

export type ConvertorRows2TableDataOpts = {
  chunkSize?: number;
  typeParser?: TypeParser;
  templateManager: ExcelTemplateManager<CellImportOptions>;
  onTrigger: (section: SheetSection, data: any | any[]) => Promise<void>;
  onErrors: (err: any[]) => Promise<void>;
};

export class ConvertorRows2TableData {
  private chunkSize: number;
  private typeParser: TypeParser;
  private templateManager: ExcelTemplateManager<CellImportOptions>;
  private tableCols: string[] = [];
  private container: TableData<any> = { header: {}, footer: {}, table: [] };
  private section: SheetSection = "header";
  private addresses: TableData<any> = { header: {}, footer: {}, table: [] };

  onTrigger: (section: SheetSection, data: any | any[]) => Promise<void> = async () => {};
  onErrors: (err: any) => Promise<void> = async () => {};

  constructor(opts: ConvertorRows2TableDataOpts) {
    this.chunkSize = opts.chunkSize ?? 10;
    this.typeParser = opts.typeParser ?? new TypeParser();
    this.templateManager = opts.templateManager;
    this.onErrors = opts?.onErrors ?? this.onErrors;
    this.onTrigger = opts?.onTrigger ?? this.onTrigger;
    this.tableCols = this.templateManager.GroupCells.table.map((e) => e.keyName);
  }

  async push(row: exceljs.Row | null) {
    if (row === null) return await this.complete();

    const section = this.getSection(row.number);

    let values: any = {};
    let address: any = {};

    for (let i = 0; i < row.cellCount; i++) {
      const cell = row.getCell(i + 1);
      const keyName = this.getKeyField(cell, section);
      if (!keyName) continue;
      values[keyName] = cell.value as any;
      address[keyName] = cell.address;
    }
    if (Object.keys(values).length > 0) this.addContainer(section, values, address);
    const isTrigger = this.isTrigger(section, row.number);
    if (isTrigger) {
      const data = await this.pop(section);
      if (data) await this.onTrigger(section, data);
    }
    this.section = section;
  }

  private async pop(section: SheetSection) {
    const data = Array.isArray(this.container[section]) ? this.container[section] : [this.container[section]];
    const address = Array.isArray(this.addresses[section]) ? this.addresses[section] : [this.addresses[section]];
    const error = [];
    if (section === "table") {
      this.container.table = [];
      this.addresses.table = [];
    } else {
      this.container[section] = {};
      this.addresses[section] = {};
    }
    for (let i = 0; i < data.length; i++) {
      const { errors, row } = this.formatValue(data[i], address[i], section);
      if (row) data[i] = row;
      if (errors && errors.length > 0) error.push(...errors);
    }
    if (error.length !== 0) await this.onErrors(error);

    return !data[0] || Object.keys(data[0]).length === 0 ? null : section === "table" ? data : data[0];
  }

  private getSection(rowIndex: number) {
    const beginTable = this.templateManager.ActualTableStartRow ?? DEFAULT_BEGIN_TABLE;
    const endTable = this.templateManager.ActualTableEndRow ?? DEFAULT_END_TABLE;
    const section = ReaderExceljsHelper.getSection(rowIndex, beginTable, endTable);

    return section;
  }

  private isTrigger(section: SheetSection, rowIndex: number) {
    const beginTable = this.templateManager.ActualTableStartRow ?? DEFAULT_BEGIN_TABLE;
    const chunkSize = this.container?.table?.length;

    if (beginTable === rowIndex) return true;
    else if (section === "footer" && this.section === "table") return true;
    else if (this.chunkSize <= chunkSize && section === "table") return true;
    return false;
  }

  private addContainer(section: SheetSection, data: any, address: any) {
    if (!this.addresses[section] && section === "table") this.addresses.table = [];
    if (!this.addresses[section] && section !== "table") this.addresses[section] = {};

    if (section === "table") {
      this.container.table.push(data);
      this.addresses.table.push(address);
    } else {
      this.container[section] = { ...this.container[section], ...data };
      this.addresses[section] = { ...this.addresses[section], ...address };
    }
  }

  private formatValue(data: any | any[], address: any, section: SheetSection) {
    const row = { ...data };
    const errors: ValidateImportError[] = [];
    if (!this.templateManager.GroupCells[section]) return { row: null, errors: null };
    this.templateManager.GroupCells[section]?.forEach((cell) => {
      let value = row[cell.keyName];
      if (cell.setValue) value = cell.setValue(value, row);
      if (cell.type && cell.type !== "virtual") value = (this.typeParser as any)[cell.type](value);

      try {
        validateCellImport(value, cell, address[cell.keyName], cell.keyName);
        row[cell.keyName] = value;
      } catch (error) {
        errors.push(error as any);
      }
    });

    return { row, errors };
  }

  private getKeyField(cell: exceljs.Cell, section: SheetSection) {
    if (section === "table") {
      return this.tableCols[cell.fullAddress.col - 1];
    }
    if (section === "header" && this.templateManager.GroupCells.header) {
      const headerCells = this.templateManager.GroupCells.header;
      const index = headerCells.findIndex((e) => e.address === cell.address);
      if (index >= 0) return headerCells[index].keyName;
      else return "";
    }
    if (this.templateManager.GroupCells.footer) {
      const footerCells = this.templateManager.GroupCells.footer;
      const index = footerCells.findIndex((e) => {
        const rownum = (this.templateManager.ActualTableEndRow ?? DEFAULT_END_TABLE) + (e.fullAddress?.row ?? 0);
        return cell.address === `${e.fullAddress?.col}${rownum}`;
      });
      if (index >= 0) return footerCells[index].keyName;
      else return "";
    }
    return null;
  }

  private async complete() {
    let data = await this.pop(this.section);
    if (data) await this.onTrigger(this.section, data);

    data = await this.pop("footer");
    if (data) await this.onTrigger("footer", data);

    await this.onTrigger(this.section, null);
  }
}
