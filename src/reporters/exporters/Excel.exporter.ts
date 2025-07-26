import * as exceljs from "exceljs";
import * as fs from "fs";
import { Writable } from "stream";
import { EventRegister } from "../../common/core/ListEvents.js";
import { ExcelTemplateManager } from "../../common/core/Template.js";
import { TableData } from "../../common/types/common-type.js";
import { CellReportOptions } from "../../common/types/report-template.type.js";
import { PartialDataTransfer, PartialDataTransferRunner } from "../PartialDataTransferV2.js";
import { IExporter } from "./IExporter.js";
import { ExcelProcessor, ExcelStreamProcessor } from "./proccessor/ExcelProcessor.js";
import { RingPromise } from "../../common/core/RingPromise.js";
import { DataTranfer, PartialDataHandler } from "../IPartialDataHandler.js";

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
  private excelProcessor: ExcelProcessor = {} as any;

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
    const workBook = new exceljs.Workbook();
    this.Event.emitEvent("onFile");
    // await this.execute(workBook, data);
    await workBook.xlsx.writeFile(reportPath);
  }

  async toBuffer(data: TableData, opts?: ExcelExporterOptions): Promise<Buffer>;
  async toBuffer(data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): Promise<Buffer>;
  async toBuffer(data: TableData | TableDataPartialDataTransfer, opts?: ExcelExporterOptions): Promise<Buffer> {
    const workBook = new exceljs.Workbook();
    this.Event.emitEvent("onFile");

    return (await workBook.xlsx.writeBuffer()) as unknown as Buffer;
  }

  streamTo(filePath: string, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): void;
  streamTo(stream: Writable, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): void;
  streamTo(arg1: unknown, data: TableDataPartialDataTransfer, opts?: ExcelExporterOptions): void {
    this.opts = { style: opts?.style ?? "use-style" };
    const stream = arg1 instanceof Writable ? arg1 : fs.createWriteStream(arg1 as string);

    const workBook = new exceljs.stream.xlsx.WorkbookWriter({
      stream,
      useSharedStrings: opts?.useSharedStrings,
      useStyles: opts?.style === "use-style" ? true : false,
      zip: opts?.zip,
    });
    this.Event.emitEvent("onFile");

    this.execute(workBook, data, ExcelStreamProcessor);
  }

  private async execute(
    workBook: exceljs.Workbook,
    data: TableDataPartialDataTransfer,
    classProcessor: typeof ExcelStreamProcessor | typeof ExcelProcessor = ExcelProcessor
  ): Promise<void> {
    this.excelProcessor = new classProcessor({
      workBook,
      template: this.Template,
      event: this.event,
      header: data.header,
      footer: data.footer,
      style: this.opts.style,
    });
    const originalSheetName = this.Template.SheetInformation.sheetName;
    const tableData = data.table as unknown as PartialDataTransferRunner;

    const partialDataHandler = new PartialDataHandler(originalSheetName, this.createTask());

    await tableData.init(partialDataHandler, originalSheetName);
    this.Event.emitEvent("start");
    await tableData.start();
  }

  private createTask() {
    return async (args: DataTranfer) => {
      this.excelProcessor.pushData(args.sheetName ?? "", args.items, args.sheetCompleted);
      if (this.excelProcessor instanceof ExcelStreamProcessor && args.isCompleted) {
        await this.excelProcessor.finalizeWorkbook();
      }
    };
  }
}
