import { FilterImportHandler } from "../common/types/importer.type.js";

export abstract class ImporterHandler {
  async run(data: any, filter: FilterImportHandler) {
    if (this.filter(filter)) return await this.handle(data);
    return data;
  }

  filter(filterOtps: FilterImportHandler): boolean {
    return true;
  }

  abstract handle(result: ImporterHandler): Promise<any>;
}
