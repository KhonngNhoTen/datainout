import { SheetSection } from "../type";

export type FilterImportHandler = {
  sheetIndex: number;
  section: SheetSection;
};

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
