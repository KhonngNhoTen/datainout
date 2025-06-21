import * as fs from "fs";
import { Readable } from "stream";
import { pathImport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
import { BaseReaderStream } from "./readers/BaserReaderStream.js";
import { ImporterBaseReaderStreamType, ImporterBaseReaderType, ImporterHandlerFunction } from "../common/types/importer.type.js";
import { BaseReader } from "./readers/BaseReader.js";
import { ExcelJsReader } from "./readers/exceljs/ExcelJsReader.js";
import { ExcelJsCsvReader } from "./readers/csv/ExceljsCsvReader.js";
import { ExcelJsStreamReader } from "./readers/exceljs/ExcelJsStreamReader.js";
import { IBaseStream } from "../common/core/ListEvents.js";

export class Importer {
  protected templatePath: string;
  constructor(templatePath: string) {
    this.templatePath = pathImport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
  }

  async load(filePath: string, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderType, chunkSize?: number): Promise<any>;
  async load(buffer: Buffer, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderType, chunkSize?: number): Promise<any>;
  async load(arg: unknown, handlers: ImporterHandlerFunction[], type: ImporterBaseReaderType = "excel", chunkSize?: number) {
    if (!type) type = "excel";
    const reader = this.createBaseReader(type);
    await reader.run(this.templatePath, arg as any, handlers, chunkSize);
  }

  createStream(arg: string, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): IBaseStream;
  createStream(arg: Readable, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): IBaseStream;
  createStream(arg: unknown, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): IBaseStream {
    if (!type) type = "excel-stream";
    const fsStream = typeof arg === "string" ? fs.createReadStream(pathImport(arg, "excelSampleDir")) : arg;
    const readerStream = new ExcelJsStreamReader(this.templatePath, fsStream as any, handlers);
    return readerStream as IBaseStream;
  }

  private createBaseReader(type: string) {
    if (type === "excel") return new ExcelJsReader();
    if (type === "csv") return new ExcelJsCsvReader();
    throw new Error(`Type ${type} not supports`);
  }
}
