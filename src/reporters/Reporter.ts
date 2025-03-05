import { CronTime } from "cron";
import { Exporter } from "./exporters/Exporter.js";
import { HandlerCron } from "./schedule/Cron.js";
import { CronManager } from "./schedule/CronManager.js";
import { ReportData, ExporterList, ExporterFactory, CreateStreamOpts } from "./type.js";
import { pathReport } from "../helper/path-file.js";
import { exporterFactory } from "./exporters/ExporterFactory.js";
import { getConfig } from "../datainout-config.js";
import { PassThrough } from "stream";
import { ReportDataIterator } from "./ReportDataIterator.js";

type ReporterOptions = {
  exporterType: ExporterList;
  templatePath: string;
};

export class Reporter {
  private exporter: Exporter;
  private cronManager: CronManager;

  constructor(opts: ReporterOptions) {
    const myExporterFactory: ExporterFactory = getConfig()?.report?.expoterFactory ?? exporterFactory;
    this.exporter = myExporterFactory(opts.exporterType, pathReport(opts.templatePath, "templateDir"));
    this.cronManager = new CronManager(this);
  }
  async writeFile(data: ReportData, reportPath: string): Promise<any>;
  async writeFile(data: ReportData[], reportPath: string): Promise<any>;
  async writeFile(data: ReportData | ReportData[], reportPath: string) {
    console.log(`Gernerating report ....`);
    await this.exporter.writeFile(data, reportPath);
    console.log(`Generate report successfully. File report at ${reportPath}`);
  }

  async createBuffer(data: ReportData): Promise<Buffer>;
  async createBuffer(data: ReportData[]): Promise<Buffer>;
  async createBuffer(data: ReportData | ReportData[]): Promise<Buffer> {
    return await this.exporter.buffer(data);
  }

  crons() {
    return this.crons;
  }

  createCron(scheduling: string | Date, name: string, onTick?: HandlerCron, cronTime?: CronTime) {
    return this.cronManager.createCron(scheduling, name, onTick, cronTime);
  }

  createStream(opts: CreateStreamOpts | (Omit<CreateStreamOpts, "data"> & { data: ReportDataIterator[] })): PassThrough {
    if (opts.data[0] instanceof ReportDataIterator) opts.data = opts.data.map((e) => ({ table: e })) as any;

    const options: CreateStreamOpts = opts as any;
    const writerStream = this.exporter.writerStream(options);

    let countDoneData = 0;
    for (let i = 0; i < options.data.length; i++) {
      const data = options.data[i];
      writerStream.setContent({ sheetName: data.sheetName, footer: data.header, header: data.header });

      const dataIteratorStream = data.table.createStream();

      dataIteratorStream.on("data", async (data) => writerStream.add(JSON.parse(data.toString()), i));

      dataIteratorStream.on("end", async () => {
        await writerStream.doneSheet(i);
        countDoneData++;
        if (countDoneData === options.data.length) await writerStream.allDone();
      });
    }
    return writerStream.stream();
  }
}
