import * as fs from "fs";
import { pathImport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
import { Readable } from "stream";
import { BaseReaderStream } from "./readers/BaserReaderStream.js";
import { ImporterBaseReaderStreamType, ImporterBaseReaderType, ImporterHandlerFunction } from "../common/types/importer.type.js";
import { ReaderContainer } from "./readers/ReaderFactory.js";

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
    const reader = ReaderContainer.get(type).reader;
    await reader.run(this.templatePath, arg as any, handlers, chunkSize);
  }

  async createStream(arg: string, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): Promise<BaseReaderStream>;
  async createStream(arg: Readable, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): Promise<BaseReaderStream>;
  async createStream(arg: unknown, handlers: ImporterHandlerFunction[], type?: ImporterBaseReaderStreamType): Promise<BaseReaderStream> {
    if (!type) type = "excel-stream";
    const fsStream = typeof arg === "string" ? fs.createReadStream(pathImport(arg, "excelSampleDir")) : arg;
    const readerStream = ReaderContainer.get(type).reader;
    await readerStream.run(this.templatePath, fsStream, handlers, 10);
    return readerStream as BaseReaderStream;
  }
}
