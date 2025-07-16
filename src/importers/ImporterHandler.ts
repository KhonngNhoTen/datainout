import { FilterImportHandler } from "../common/types/importer.type.js";

export abstract class ImporterHandler<T> {
  protected eachRow: boolean = false;
  constructor(eachRow: boolean = false) {
    this.eachRow = eachRow;
  }

  async run(tabledata: any, filterImporter: FilterImportHandler) {
    try {
      if (tabledata instanceof Error) await this.catch(tabledata);
      else if (tabledata.header) await this.handleHeader(tabledata.header, filterImporter);
      else if (tabledata.footer) await this.handleHeader(tabledata.footer, filterImporter);
      else if (tabledata.table && !this.eachRow) await this.handleChunk(tabledata.table, filterImporter);
      else if (tabledata.table && this.eachRow)
        for (let i = 0; i < tabledata.table.length; i++)
          try {
            await this.handleRow(tabledata.table[i], filterImporter);
          } catch (error) {
            await this.catch(error as any);
          }
    } catch (e) {
      await this.catch(e as any);
    }
  }

  protected async catch(error: Error): Promise<void> {}
  protected async handleChunk(chunk: T[], filter: FilterImportHandler): Promise<void> {}
  protected async handleRow(data: T, filter: FilterImportHandler): Promise<void> {}
  protected async handleHeader(header: any, filter: FilterImportHandler): Promise<void> {}
  protected async handleFooter(footer: any, filter: FilterImportHandler): Promise<void> {}
}
