import { getFileExtension } from "../../helpers/get-file-extension.js";
import { sortByAddress } from "../../helpers/sort-by-address.js";
import { BaseAttribute, SheetExcelOption, SheetSection, TableExcelOptions } from "../types/common-type.js";

export interface IExcelTemplateManager<T extends BaseAttribute> {
  get<K>(index: number): K;
  get<K>(key: string): K;

  add(cell: T): void;
  add(cell: T[]): void;

  update(index: number, cell: T): void;
  update(key: string, cell: T): void;

  remove(index: number): void;
  remove(key: string): void;
}

export class ExcelTemplateManager<T extends BaseAttribute> implements IExcelTemplateManager<T> {
  private actualTableStartRow?: number;
  private actualTableEndRow?: number | undefined;

  private sheets: SheetExcelOption<T>[] = [];
  private currentSheetIndex = 0;
  private groupCells: { [k in SheetSection]: T[] } = {} as any;
  private isNullTemplate: boolean = false;

  constructor(templatePath?: string) {
    this.isNullTemplate = !templatePath;
    this.sheets = templatePath ? this.getTemplate(templatePath) : [];
  }

  public get ActualTableStartRow(): number | undefined {
    return this.actualTableStartRow;
  }

  public get ActualTableEndRow(): number | undefined {
    return this.actualTableEndRow;
  }

  public defineActualTableStartRow(actualTableStartRow?: number) {
    if (!actualTableStartRow || actualTableStartRow <= 0) return;
    if (this.ActualTableStartRow) return;
    this.actualTableStartRow = actualTableStartRow;
  }

  public defineActualTableEndRow(actualTableEndRow?: number) {
    if (!actualTableEndRow || actualTableEndRow <= 0) return;
    if (this.ActualTableEndRow) return;
    this.actualTableEndRow = actualTableEndRow;
  }

  public set SheetIndex(sheetIndex: number) {
    if (sheetIndex < 0) return;
    this.currentSheetIndex = sheetIndex;
    this.groupCells = this.formatSheet();
  }

  public get SheetInformation(): Omit<SheetExcelOption<T>, "cells"> {
    return this.SheetTemplate;
  }

  public get SheetTemplate() {
    return this.sheets[this.currentSheetIndex];
  }

  public get Sheets() {
    return this.sheets;
  }

  public get GroupCells() {
    return this.groupCells;
  }

  get<K>(index: number): K;
  get<K>(key: string): K;
  get<K>(arg: unknown): K {
    if (this.isNullTemplate) throw new Error("Template is null. Please check template path");
    let index = 0;
    if (typeof arg === "string") index = this.findIndexByKeyName(arg);
    else if (typeof arg === "number") index = arg;
    return this.SheetTemplate.cells[index] as any;
  }

  add(cell: T): void;
  add(cell: T[]): void;
  add(cell: unknown): void {
    if (this.isNullTemplate) throw new Error("Template is null. Please check template path");
    if (!Array.isArray(cell)) cell = [cell];
    this.SheetTemplate.cells.push(...(cell as any[]));
  }

  update(index: number, cell: T): void;
  update(key: string, cell: T): void;
  update(key: unknown, cell: T) {
    if (this.isNullTemplate) throw new Error("Template is null. Please check template path");
    let index = 0;
    if (typeof key === "string") index = this.findIndexByKeyName(key);
    else if (typeof key === "number") index = key;
    this.SheetTemplate.cells[index] = cell;
  }

  remove(index: number): void;
  remove(key: string): void;
  remove(key: unknown) {
    if (this.isNullTemplate) throw new Error("Template is null. Please check template path");
    let index = 0;
    if (typeof key === "string") index = this.findIndexByKeyName(key);
    else if (typeof key === "number") index = key;

    this.SheetTemplate.cells.splice(index, 1);
  }

  private findIndexByKeyName(key: string) {
    const index = this.SheetTemplate.cells.findIndex((e) => e.keyName === key);
    if (index < 0) throw new Error("Not found template with key: " + key);
    return index;
  }

  private getTemplate(templatePath: string) {
    const template: TableExcelOptions<SheetExcelOption<T>> =
      getFileExtension(templatePath) === "js" ? require(templatePath) : require(templatePath).default;
    return template.sheets;
  }

  protected formatSheet() {
    const defaultGroup: { header: T[]; table: T[]; footer: T[] } = {} as any;
    if (this.isNullTemplate) return defaultGroup;

    const excel: any = this.SheetTemplate?.cells.reduce((acc, cell) => {
      if (!acc[cell.section]) acc[cell.section] = [cell];
      else acc[cell.section]?.push(cell);
      return acc;
    }, {} as typeof defaultGroup);
    const keys = Object.keys(excel);
    for (let i = 0; i < keys.length; i++) excel[keys[i]] = sortByAddress(excel[keys[i]]);

    return excel;
  }
}
