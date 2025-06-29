import * as fs from "fs";
import { Readable } from "stream";
import { pathImport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
import { ImporterBaseReaderStreamType, ImporterHandlerInstance, ImporterLoadFunctionOpions } from "../common/types/importer.type.js";
import { ExcelJsReader } from "./readers/exceljs/ExcelJsReader.js";
import { ExcelJsCsvReader } from "./readers/csv/ExceljsCsvReader.js";
import { ExcelJsStreamReader } from "./readers/exceljs/ExcelJsStreamReader.js";
import { IBaseStream } from "../common/core/ListEvents.js";
import { CellImportOptions, SheetImportOptions, TableImportOptions } from "../common/types/import-template.type.js";
import { getFileExtension } from "../helpers/get-file-extension.js";

export class Importer {
  protected templatePath: string;
  protected templates: SheetImportOptions[];
  constructor(templatePath: string) {
    this.templatePath = pathImport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
    this.templates = this.getTemplates(this.templatePath).sheets;
  }

  async load(filePath: string, handler: ImporterHandlerInstance, opts?: ImporterLoadFunctionOpions): Promise<any>;
  async load(buffer: Buffer, handler: ImporterHandlerInstance, opts?: ImporterLoadFunctionOpions): Promise<any>;
  async load(arg: unknown, handler: ImporterHandlerInstance, opts?: ImporterLoadFunctionOpions) {
    const type = opts?.type ?? "excel";
    const reader = this.createBaseReader(type);
    await reader.run(this.templates, arg as any, handler, opts);
  }

  createStream(arg: string, handler: ImporterHandlerInstance, type?: ImporterBaseReaderStreamType): Omit<IBaseStream, "onError">;
  createStream(arg: Readable, handler: ImporterHandlerInstance, type?: ImporterBaseReaderStreamType): Omit<IBaseStream, "onError">;
  createStream(arg: unknown, handler: ImporterHandlerInstance, type?: ImporterBaseReaderStreamType): Omit<IBaseStream, "onError"> {
    if (!type) type = "excel-stream";
    const fsStream = typeof arg === "string" ? fs.createReadStream(pathImport(arg, "excelSampleDir")) : arg;
    const readerStream = new ExcelJsStreamReader(this.templates, fsStream as any, handler);
    return readerStream as Omit<IBaseStream, "onError">;
  }

  private createBaseReader(type: string) {
    if (type === "excel") return new ExcelJsReader();
    if (type === "csv") return new ExcelJsCsvReader();
    throw new Error(`Type ${type} not supports`);
  }

  public addCellTemplate(cell: CellImportOptions, sheetIndex?: number): void;
  public addCellTemplate(cell: CellImportOptions[], sheetIndex?: number): void;
  public addCellTemplate(arg: unknown, sheetIndex?: number) {
    sheetIndex = sheetIndex ?? 0;
    const cells: CellImportOptions[] = Array.isArray(arg) ? arg : [arg];
    this.templates[sheetIndex].cells.push(...cells);
  }

  public getCellTemplate(key: string, sheetIndex: number = 0): null | CellImportOptions {
    const index = this.templates[sheetIndex].cells.findIndex((e) => (e.keyName = key));
    if (index < 0) return null;
    return this.templates[sheetIndex].cells[index];
  }

  public setCellTemplate(cell: CellImportOptions, sheetIndex: number = 0) {
    const index = this.templates[sheetIndex].cells.findIndex((e) => (e.keyName = cell.keyName));
    if (index < 0) return false;
    this.templates[sheetIndex].cells[index] = cell;
    return true;
  }

  protected getTemplates(templatePath: string): TableImportOptions {
    return getFileExtension(templatePath) === "js" ? require(templatePath) : require(templatePath).default;
  }
}
