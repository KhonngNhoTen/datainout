import { TypeParser } from "../../../helpers/parse-type.js";
import { BaseReader } from "../BaseReader.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";
import { ImporterHandler } from "../../ImporterHandler.js";
import { FilterImportHandler } from "../../../common/types/importer.type.js";
import { SheetSection } from "../../../common/types/common-type.js";
import { ConvertorRows2TableData } from "../../../helpers/convert-row-to-table-data-v2.js";

export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ReaderExceljsHelper = null as any;

  constructor(templateManager: ExcelTemplateManager<CellImportOptions>) {
    super({ type: "excel", typeParser: new TypeParser(), templateManager });
  }

  async load(arg: unknown): Promise<any> {
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: async () => {
        await this.convertorRows2TableData.push(null);
      },
      onRow: async (data) => {
        await this.convertorRows2TableData.push(data.detail);
      },
      isSampleExcel: false,
      templateManager: this.templateManager,
    });

    this.convertorRows2TableData = new ConvertorRows2TableData({
      onTrigger: async (sect, data) => await this.callHandler(sect, data),
      onErrors: async (err) => await this.onErrors(err),
      chunkSize: this.options?.chunkSize,
      templateManager: this.templateManager,
    });

    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }
}
