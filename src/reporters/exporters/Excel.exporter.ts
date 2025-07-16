import * as exceljs from "exceljs";
import * as fs from "fs";
import { Writable } from "stream";
import { EventRegister } from "../../common/core/ListEvents.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { TableData } from "../../common/types/common-type.js";
import { CellReportOptions } from "../../common/types/report-template.type.js";
import { PartialDataTransfer } from "../PartialDataTransfer.js";
import { IExporter } from "./IExporter.js";
import { ExcelProcessor, ExcelStreamProcessor } from "./proccessor/ExcelProcessor.js";
import { PromiseBag } from "../../common/core/PromiseBag.js";

type ExcelExporterOptions = {
  useSharedStrings?: boolean;
  zip?: Partial<exceljs.stream.xlsx.ArchiverZipOptions>;
  style?: "no-style" | "no-style-no-header" | "use-style";
};

type TableDataPartialDataTransfer = {
  header?: any;
  footer?: any;
  table: PartialDataTransfer;
};

export class ExcelExporter implements IExporter {
  private templatePath: string;
  private template: ExcelTemplateManager<CellReportOptions> = {} as any;
  private event: EventRegister = new EventRegister();
  private opts: { style: "no-style" | "no-style-no-header" | "use-style" } = { style: "use-style" };

  constructor(templatePath: string) {
    this.templatePath = templatePath;
    this.template = new ExcelTemplateManager(this.templatePath);
    this.template.SheetIndex = 0;
  }

  public get Template(): ExcelTemplateManager<CellReportOptions> {
    return this.template;
  }

  public get Event(): EventRegister {
    return this.event;
  }

  async write(reportPath: string, data: TableData, opts?: ExcelExporterOptions): Promise<void>;
  async write(reportPath: string, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): Promise<void>;
  async write(reportPath: string, data: TableData | TableDataPartialDataTransfer, opts?: ExcelExporterOptions): Promise<void> {
    this.Event.emitEvent("onFile");
    const workBook = new exceljs.Workbook();
    await this.execute(workBook, data);
    await workBook.xlsx.writeFile(reportPath);
  }

  async toBuffer(data: TableData, opts?: ExcelExporterOptions): Promise<Buffer>;
  async toBuffer(data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): Promise<Buffer>;
  async toBuffer(data: TableData | TableDataPartialDataTransfer, opts?: ExcelExporterOptions): Promise<Buffer> {
    this.Event.emitEvent("onFile");
    const workBook = new exceljs.Workbook();
    await this.execute(workBook, data);
    return (await workBook.xlsx.writeBuffer()) as unknown as Buffer;
  }

  streamTo(filePath: string, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): void;
  streamTo(stream: Writable, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): void;
  streamTo(arg1: unknown, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): void {
    this.Event.emitEvent("onFile");
    this.opts = { style: opts?.style ?? "use-style" };
    const stream = arg1 instanceof Writable ? arg1 : fs.createWriteStream(arg1 as string);

    const workBook = new exceljs.stream.xlsx.WorkbookWriter({
      stream,
      useSharedStrings: opts?.useSharedStrings,
      useStyles: opts?.style === "use-style" ? true : false,
      zip: opts?.zip,
    });

    this.execute(workBook, data, ExcelStreamProcessor);
  }

  private async execute(
    workBook: exceljs.Workbook,
    data: TableData | TableDataPartialDataTransfer,
    classProcessor: typeof ExcelStreamProcessor | typeof ExcelProcessor = ExcelProcessor
  ): Promise<void> {
    const excelProcessor = new classProcessor({
      workBook,
      template: this.Template,
      event: this.event,
      header: data.header,
      footer: data.footer,
      style: this.opts.style,
    });
    const originalSheetName = this.Template.SheetInformation.sheetName;

    if (!(data.table instanceof PartialDataTransfer)) {
      const tableData = (data as TableData).table;
      for (let i = 0; tableData && i < tableData.length; i++) {
        excelProcessor.pushData(originalSheetName, tableData);
      }
    } else {
      const tableData = data.table as PartialDataTransfer;
      await tableData.run(originalSheetName, async (args) => {
        excelProcessor.pushData(args.sheetName, args.items, args.sheetStatus === "completed");
        if (excelProcessor instanceof ExcelStreamProcessor && args.status === "completed") await excelProcessor.finalizeWorkbook();
      });
    }
  }

  // private createPromise(callBack: (...args: any[]) => {}, args: any): Promise<void> {
  //   return new Promise((res, rej) => {
  //     callBack(args.sheetName, args.items, args.sheetStatus === "completed");
  //     res();
  //   });
  // }
}
