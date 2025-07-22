import { CellImportOptions } from "../common/types/import-template.type.js";
import { TypeParser } from "./parse-type.js";
import { ConvertorRows2TableDataOpts } from "../common/types/convert-row-to-table-data.type.js";
import { ExcelTemplateManager } from "../common/core/Template.js";

export class ConvertorRows2TableData2 {
  private templateManager: ExcelTemplateManager<CellImportOptions>;
  private typeParser: TypeParser;
  private batchSize: number;
  private headerCols: { [col: number]: string }[] = [];
  private footerCols: { [col: number]: string } = {};
  private tableCols: { [col: number]: string } = {};

  constructor(opts?: ConvertorRows2TableDataOpts) {
    this.batchSize = opts?.chunkSize ?? 10;
    this.typeParser = opts?.typeParser ?? new TypeParser();
    this.templateManager = opts?.templateManager ?? new ExcelTemplateManager();
    this.templateManager.GroupCells.table.forEach((e) => {
      this.tableCols[e.fullAddress?.col ?? 0] = e.keyName;
    });
    this.templateManager.GroupCells.footer.forEach((e) => {
      this.footerCols[e.fullAddress?.col ?? 0] = e.keyName;
    });
    this.templateManager.GroupCells.header.forEach((e) => {
      this.headerCols[e.fullAddress?.col ?? 0] = e.keyName;
    });
  }

  private convertHeader(data: any[]) {
    const _header = this.headerCols[0];
    const result: any = {};
    Object.keys(this.headerCols).forEach((e) => (result[_header[+e]] = data[+e]));
    return result;
  }
  private convertFooter(data: any[]) {
    const _footer = this.footerCols[0];
    const result: any = {};
    Object.keys(this.footerCols).forEach((e) => (result[_footer[+e]] = data[+e]));
    return result;
  }
  private convertTable(data: any[]) {
    return true;
  }
}
