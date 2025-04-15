import { pathImport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
import { Readable } from "stream";
import { BaseReaderStream } from "./readers/BaserReaderStream.js";
import { ImporterBaseReaderStreamType, ImporterBaseReaderType } from "../common/types/importer.type.js";
import { ReaderContainer } from "./readers/ReaderFactory.js";

export class Importer {
  protected templatePath: string;
  constructor(templatePath: string) {
    this.templatePath = pathImport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
  }

  async load(filePath: string, type?: ImporterBaseReaderType, chunkSize?: number): Promise<any>;
  async load(buffer: Buffer, type?: ImporterBaseReaderType, chunkSize?: number): Promise<any>;
  async load(arg: unknown, type: ImporterBaseReaderType = "excel", chunkSize?: number) {
    if (!type) type = "excel";
    const reader = ReaderContainer.get(type).reader;
    await reader.run(this.templatePath, arg as any, [], chunkSize);
  }

  createStream(reable: Readable, type?: ImporterBaseReaderStreamType): BaseReaderStream {
    if (!type) type = "excel-stream";
    const readerStream = ReaderContainer.get(type).reader;
    readerStream.run(this.templatePath, reable, [], 10);
    return readerStream as BaseReaderStream;
  }
}
