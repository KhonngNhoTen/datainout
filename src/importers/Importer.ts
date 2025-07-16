import * as fs from "fs";
import { Readable } from "stream";
import { pathImport } from "../helpers/path-file.js";
import { getConfig } from "../helpers/datainout-config.js";
import { ImporterBaseReaderStreamType, ImporterHandlerInstance, ImporterLoadFunctionOpions } from "../common/types/importer.type.js";
import { ExcelJsReader } from "./readers/exceljs/ExcelJsReader.js";
import { ExcelJsCsvReader } from "./readers/csv/ExceljsCsvReader.js";
import { ExcelJsStreamReader } from "./readers/exceljs/ExcelJsStreamReader.js";
import { IBaseStream } from "../common/core/ListEvents.js";
import { CellImportOptions } from "../common/types/import-template.type.js";
import { ExcelTemplateManager, IExcelTemplateManager } from "../common/core/Template.js";
import { Piscina } from "piscina";

export class Importer {
  protected templatePath: string;
  protected excelsTemplate: ExcelTemplateManager<CellImportOptions>;
  protected workerPools?: Piscina;

  constructor(templatePath: string) {
    this.templatePath = pathImport(templatePath, "templateDir");
    this.templatePath = `${this.templatePath}${getConfig().templateExtension ?? ".js"}`;
    this.excelsTemplate = new ExcelTemplateManager(this.templatePath);
  }

  public get ExcelTemplate(): IExcelTemplateManager<CellImportOptions> {
    return this.excelsTemplate;
  }

  async load(filePath: string, handler: ImporterHandlerInstance, opts?: ImporterLoadFunctionOpions): Promise<any>;
  async load(buffer: Buffer, handler: ImporterHandlerInstance, opts?: ImporterLoadFunctionOpions): Promise<any>;
  async load(arg: unknown, handler: ImporterHandlerInstance, opts?: ImporterLoadFunctionOpions) {
    const type = opts?.type ?? "excel";
    const reader = this.createBaseReader(type);
    await reader.run(this.excelsTemplate, arg as any, handler, opts);
  }

  createStream(arg: string, handler: ImporterHandlerInstance, type?: ImporterBaseReaderStreamType): Omit<IBaseStream, "onError">;
  createStream(arg: Readable, handler: ImporterHandlerInstance, type?: ImporterBaseReaderStreamType): Omit<IBaseStream, "onError">;
  createStream(arg: unknown, handler: ImporterHandlerInstance, type?: ImporterBaseReaderStreamType): Omit<IBaseStream, "onError"> {
    if (!type) type = "excel-stream";
    const fsStream = typeof arg === "string" ? fs.createReadStream(pathImport(arg, "excelSampleDir")) : arg;
    //   if (opts?.workerSize) this.workerPools = createWorkerPool(opts?.workerSize);

    const readerStream = new ExcelJsStreamReader(this.excelsTemplate, fsStream as any, handler);
    return readerStream as unknown as Omit<IBaseStream, "onError">;
  }

  private createBaseReader(type: string) {
    if (type === "excel") return new ExcelJsReader();
    if (type === "csv") return new ExcelJsCsvReader();
    throw new Error(`Type ${type} not supports`);
  }
}
