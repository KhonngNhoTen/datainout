import * as exceljs from "exceljs";
import { EventType, SheetSection, TableData } from "../common/types/common-type.js";
import { RowDataHelper } from "../common/types/excel-reader-helper.type.js";
import { CellImportOptions, SheetImportOptions } from "../common/types/import-template.type.js";
import { TypeParser } from "./parse-type.js";
import { DEFAULT_BEGIN_TABLE, DEFAULT_END_TABLE, ReaderExceljsHelper } from "./excel.helper.js";
import { validateCellImport } from "./validate-cell-importer.js";
import { ConvertorRows2TableDataOpts, GroupValueRow } from "../common/types/convert-row-to-table-data.type.js";
import { ValidateImportError } from "../common/error/ValidateError.js";
import { ExcelTemplateManager } from "../common/core/Template.js";

type PushResultTableData = {
  isTrigger: boolean;
  triggerSection: SheetSection;
  errors: ValidateImportError[];
  hasError: boolean;
};
export class ConvertorRows2TableData {
  private chunkSize: number;
  private section?: SheetSection;
  private typeParser: TypeParser;
  private container: TableData = { header: {}, footer: {}, table: [] };
  private beginTable: number = DEFAULT_BEGIN_TABLE;
  private endTable: number = DEFAULT_END_TABLE;
  private templateManager: ExcelTemplateManager<CellImportOptions>;

  constructor(opts?: ConvertorRows2TableDataOpts) {
    this.chunkSize = opts?.chunkSize ?? 10;
    this.typeParser = opts?.typeParser ?? new TypeParser();
    this.templateManager = opts?.templateManager ?? new ExcelTemplateManager();
  }

  push(row: RowDataHelper | null, template: SheetImportOptions): PushResultTableData;
  push(row: exceljs.Row | null, template: SheetImportOptions): PushResultTableData;
  push(arg: any, template: SheetImportOptions) {
    const addresses: any = {};
    // Map cell velue with cell description
    const groupValues: GroupValueRow = {};

    if (arg === null) return { isTrigger: true, triggerSection: this.section ?? "header" };

    if (ReaderExceljsHelper.isNullableRow(arg?.detail ?? arg))
      return { isTrigger: false, triggerSection: this.section ?? "header", errors: [], hasError: false };

    const { endTable, row, section } = this.getRowInformation(arg, template);
    const isTrigger = this.isTrigger(section);
    const triggerSection = this.triggerSection(section);

    if (!row) return { isTrigger: false, triggerSection: this.section ?? "header", errors: [], hasError: false };

    for (let i = 0; i < row.cellCount; i++) {
      const cell = row.getCell(i + 1);
      if (cell?.value === undefined) continue;
      const index = template.cells.findIndex((e) => this.compareByAddress(e, section, cell, row.number, endTable));
      if (index === -1) continue;
      const cellImport = template.cells[index];
      groupValues[cellImport.keyName] = cell.value;
      addresses[cellImport.keyName] = cell.address;
    }

    if (section === "table" && Object.keys(groupValues).length > 0) this.container.table?.push(groupValues);
    else if (Object.keys(groupValues).length > 0)
      this.container[section] = {
        ...this.container[section],
        ...(groupValues as any),
      };

    const { errors, hasError } = this.triggerGroupData(triggerSection, section, template, addresses, row);

    this.section = section;
    return { isTrigger, triggerSection, errors, hasError };
  }

  pushBySection(section: SheetSection, template: SheetImportOptions, rowIndex: number) {
    const result = this.formatValue(
      template.cells.filter((e: CellImportOptions) => e.section === section),
      {},
      this.typeParser,
      rowIndex
    );
    if (section === "table") this.container.table?.push(result.groupValues);
    else this.container[section] = result.groupValues;

    return {
      isTrigger: true,
      triggerSection: section,
      errors: result.errors,
      hasError: result.errors.length !== 0,
    };
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

    const beginTable = this.templateManager.ActualTableStartRow ?? DEFAULT_BEGIN_TABLE;
    const endTable = this.templateManager.ActualTableEndRow ?? DEFAULT_END_TABLE;
    if (!arg.detail) section = ReaderExceljsHelper.getSection(row, beginTable, endTable);

    return { row, section, beginTable, endTable };
  }

  private triggerGroupData(
    triggerSection: SheetSection,
    section: SheetSection,
    template: SheetImportOptions,
    addresses: any,
    row: exceljs.Row
  ) {
    const errors: ValidateImportError[] = [];
    if (triggerSection !== section) {
      const resultFormat = this.formatValue(
        template.cells.filter((e: CellImportOptions) => e.section === triggerSection),
        this.container[triggerSection],
        this.typeParser,
        row.number,
        addresses
      );
      errors.push(...resultFormat.errors);
      // Set groupvalue;
      this.container[triggerSection] = resultFormat.groupValues;
    }
    if (section === "table") {
      const groupValue = this.container.table?.pop();
      const resultFormat = this.formatValue(
        template.cells.filter((e: CellImportOptions) => e.section === section),
        groupValue,
        this.typeParser,
        row.number,
        addresses
      );
      errors.push(...resultFormat.errors);
      // Set groupvalue;
      this.container.table?.push(groupValue);
    }
    const hasError = errors.length !== 0;
    return { hasError, errors };
  }

  private formatValue(
    formattedCellImport: CellImportOptions[],
    groupValues: GroupValueRow,
    typeParser: TypeParser,
    rowIndex: number,
    addresses?: any
  ) {
    formattedCellImport = formattedCellImport.sort((a, b) => {
      const aType = a.type === "virtual" ? 2 : 1;
      const bType = b.type === "virtual" ? 2 : 1;
      return bType - aType;
    });
    const row = { ...groupValues };
    const errors: ValidateImportError[] = [];
    formattedCellImport.forEach((cell) => {
      let value = groupValues[cell.keyName];
      if (cell.setValue) value = cell.setValue(value, row);

      if (cell.type && cell.type !== "virtual") value = (typeParser as any)[cell.type](value);

      let address = addresses?.[cell.keyName] ?? undefined;

      if (!address) {
        if (cell.section === "header") address = cell.address;
        else {
          const { col } = ReaderExceljsHelper.splitAddress(cell.address ?? "");
          address = `${col}${rowIndex}`;
        }
      }

      try {
        validateCellImport(value, cell, address, cell.keyName);
        groupValues[cell.keyName] = value;
      } catch (error) {
        errors.push(error as any);
      }
    });

    return { groupValues, errors };
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
