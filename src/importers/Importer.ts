import * as fs from "fs";
import { Readable } from "stream";
import { pathImport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
import { ImporterBaseReaderStreamType, ImporterHandlerFunction, ImportFunctionOpions } from "../common/types/importer.type.js";
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

  async load(filePath: string, handlers: ImporterHandlerFunction[], opts?: ImportFunctionOpions): Promise<any>;
  async load(buffer: Buffer, handlers: ImporterHandlerFunction[], opts?: ImportFunctionOpions): Promise<any>;
  async load(arg: unknown, handlers: ImporterHandlerFunction[], opts?: ImportFunctionOpions) {
    const type = opts?.type ?? "excel";
    const reader = this.createBaseReader(type);
    await reader.run(this.templates, arg as any, handlers, opts?.chunkSize);
  }

  createStream(arg: string, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): IBaseStream;
  createStream(arg: Readable, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): IBaseStream;
  createStream(arg: unknown, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): IBaseStream {
    if (!type) type = "excel-stream";
    const fsStream = typeof arg === "string" ? fs.createReadStream(pathImport(arg, "excelSampleDir")) : arg;
    const readerStream = new ExcelJsStreamReader(this.templates, fsStream as any, handlers);
    return readerStream as IBaseStream;
  }

  private createBaseReader(type: string) {
    if (type === "excel") return new ExcelJsReader();
    if (type === "csv") return new ExcelJsCsvReader();
    throw new Error(`Type ${type} not supports`);
  }

  public addCellTemplate(cell: CellImportOptions, sheetIndex: number): void;
  public addCellTemplate(cell: CellImportOptions[], sheetIndex: number): void;
  public addCellTemplate(arg: unknown, sheetIndex: number = 0) {
    const cells: CellImportOptions[] = Array.isArray(arg) ? arg : [arg];
    this.templates[sheetIndex].cells.push(...cells);
  }

  protected getTemplates(templatePath: string): TableImportOptions {
    return getFileExtension(templatePath) === "js" ? require(templatePath) : require(templatePath).default;
  }
}
