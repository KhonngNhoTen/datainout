import { CellImportOptions, SheetImportOptions, TableImportOptions } from "../../common/types/import-template.type.js";
import { FilterImportHandler, ImporterHandlerFunction, ImporterReaderType } from "../../common/types/importer.type.js";
import { BaseReaderOptions } from "../../common/types/reader.type.js";
import { ConvertorRows2TableData } from "../../helpers/convert-row-to-table-data.js";
import { getFileExtension } from "../../helpers/get-file-extension.js";
import { TypeParser } from "../../helpers/parse-type.js";
import { sortByAddress } from "../../helpers/sort-by-address.js";

export abstract class BaseReader {
  private type: ImporterReaderType;

  protected typeParser: TypeParser;
  protected handlers: ImporterHandlerFunction[] = [];
  protected chunkSize: number = 20;
  protected convertorRows2TableData: ConvertorRows2TableData = new ConvertorRows2TableData();
  protected templates: SheetImportOptions[] = [];
  protected groupCellDescs: { header: CellImportOptions[]; table: CellImportOptions[]; footer: CellImportOptions[] } = {
    header: [],
    table: [],
    footer: [],
  };
  protected sheetIndex: number = 0;
  protected additionalTemplate: CellImportOptions[][] = [];

  constructor(opts: BaseReaderOptions) {
    this.type = opts.type;
    this.typeParser = opts.typeParser ?? new TypeParser();
  }

  protected abstract load(arg: unknown): Promise<any>;

  public async run(templates: SheetImportOptions[], arg: unknown, handlers: ImporterHandlerFunction[], opts?: any) {
    this.sheetIndex = 0;
    this.templates = templates;
    this.templates[this.sheetIndex];
    this.groupCellDescs = this.formatSheet(this.sheetIndex);
    this.chunkSize = opts?.chunkSize ?? this.chunkSize;
    this.handlers = handlers;

    await this.load(arg);
  }

  protected formatSheet(sheetIndex: number) {
    const excel: any = this.templates[sheetIndex].cells.reduce((acc, cell) => {
      if (!acc[cell.section]) acc[cell.section] = [cell];
      else acc[cell.section]?.push(cell);
      return acc;
    }, {} as { header: CellImportOptions[]; table: CellImportOptions[]; footer: CellImportOptions[] });

    const keys = Object.keys(excel);
    for (let i = 0; i < keys.length; i++) excel[keys[i]] = sortByAddress(excel[keys[i]]);

    return excel;
  }

  protected async callHandlers(data: any, filter: FilterImportHandler) {
    for (let i = 0; i < this.handlers.length; i++) {
      const handler = this.handlers[i];
      await handler(data, filter);
    }
  }

  public getType() {
    return this.type;
  }
}
