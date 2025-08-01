import { TypeParser } from "../../../helpers/parse-type.js";
import { BaseReader } from "../BaseReader.js";
import { ReaderExceljsHelper } from "../../../helpers/excel.helper.js";
import { ExcelTemplateManager } from "../../../common/core/Template.js";
import { CellImportOptions } from "../../../common/types/import-template.type.js";

export class ExcelJsReader extends BaseReader {
  private excelReaderHelper: ReaderExceljsHelper = null as any;

  constructor(templateManager: ExcelTemplateManager<CellImportOptions>) {
    super({ type: "excel", typeParser: new TypeParser(), templateManager });
  }

  async load(arg: unknown): Promise<any> {
    this.excelReaderHelper = new ReaderExceljsHelper({
      onSheet: async () => {
        if (this.globalError) throw this.globalError;
        await this.convertorRows2TableData.push(null);
      },
      onRow: async (data) => {
        if (this.globalError) throw this.globalError;
        await this.convertorRows2TableData.push(data.detail);
      },
      isSampleExcel: false,
      templateManager: this.templateManager,
    });
    const buffer = Buffer.isBuffer(arg) ? arg : Buffer.from(arg as string);
    await this.excelReaderHelper.load(buffer);
  }

  protected setGlobalError(err: Error) {
    super.setGlobalError(err);
    if (!this.options?.ignoreErrors) this.excelReaderHelper.isStop = true;
  }
}
