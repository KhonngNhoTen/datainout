import * as exceljs from "exceljs";
import { TypeParser } from "../../../helpers/parse-type.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { BaseReader } from "../BaseReader.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { RowDataHelper, SheetDataHelper } from "../../../common/types/excel-reader-helper.type.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data.js";
import { SheetSection } from "../../../common/types/common-type.js";

export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ReaderExceljsHelper = new ReaderExceljsHelper();

  constructor() {
    super({ type: "excel", typeParser: new TypeParser() });
  }

  async load(arg: unknown): Promise<any> {
    this.convertorRows2TableData = new ConvertorRows2TableData({
      chunkSize: this.options?.chunkSize,
      templateManager: this.templateManager,
    });
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: async (data) => await this.onSheet(data),
      onRow: async (data) => await this.onRow(data),
      isSampleExcel: false,
      templateManager: this.templateManager,
    });

    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }

  private async onRow(row: RowDataHelper) {
    const sheet = this.templateManager.SheetTemplate;
    await this.handleRow({ id: sheet.sheetIndex + 1, name: sheet.sheetName } as any, row.detail);
    // await this.depatchRow(
    //   async (workSheet, row) => await this.handleRow(workSheet, row),
    //   { id: sheet.sheetIndex + 1, name: sheet.sheetName },
    //   row.detail
    // );
  }

  private async onSheet(sheet: SheetDataHelper) {
    const lastestRow = sheet.lastestRow;
    const sectionIndex: any = { header: 1, table: 2, footer: 3 };
    const sections = new Set<SheetSection>();

    this.templateManager.SheetTemplate.cells.forEach((cell) => {
      if (sectionIndex[cell.section] > sectionIndex[lastestRow.section]) sections.add(cell.section);
    });

    let rowIndex = lastestRow.rowIndex;
    const arrSection = Array.from(sections);
    for (let i = 0; i < arrSection.length; i++) {
      // await this.depatchRow(
      //   async (workSheet, section, rowIndex) => await this.handleRow(workSheet, section, rowIndex),
      //   sheet.detail,
      //   arrSection[i],
      //   ++rowIndex
      // );
      await this.handleRow(sheet.detail, arrSection[i], ++rowIndex);
    }
    await this.handleRow(sheet.detail, null);
    // await this.depatchRow(async (workSheet) => await this.handleRow(workSheet, null), sheet.detail);
  }

  private async handleRow(workSheet: exceljs.Worksheet, section: SheetSection, rowIndex: number): Promise<void>;
  private async handleRow(workSheet: exceljs.Worksheet, row: exceljs.Row | null): Promise<void>;
  private async handleRow(workSheet: exceljs.Worksheet, arg: unknown, rowIndex?: number) {
    const { isTrigger, triggerSection, hasError, errors } = rowIndex
      ? this.convertorRows2TableData.pushBySection(arg as SheetSection, this.templateManager.SheetTemplate, rowIndex)
      : this.convertorRows2TableData.push(arg as any, this.templateManager.SheetTemplate);

    if (hasError) await this.handleError(errors);
    if (isTrigger) {
      const filter: FilterImportHandler = {
        section: triggerSection,
        sheetIndex: workSheet.id - 1,
        sheetName: workSheet.name,
        isHasNext: arg !== null,
      };
      const data = this.convertorRows2TableData.pop(triggerSection);
      await this.callHandlers({ [triggerSection]: data }, filter);
    }
  }

  private async handleError(errors: Error[]) {
    if (errors.length === 0) return;
    if (this.options?.ignoreErrors === true) throw errors[0];

    for (let i = 0; i < errors.length; i++) {
      await this.callHandlers(errors[i], null as any);
    }
  }
}
