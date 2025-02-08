import { CronTime } from "cron";
import { Exporter, exporterFactory } from "./exporters/Exporter";
import { HandlerCron } from "./schedule/Cron";
import { CronManager } from "./schedule/CronManager";
import { ReportData, ExporterList } from "./type";

type ReporterOptions = {
  exporterType: ExporterList;
  templatePath: string;
};

export class Reporter {
  exporter: Exporter;
  cronManager: CronManager;
  constructor(opts: ReporterOptions) {
    this.exporter = exporterFactory(opts.exporterType);
    this.exporter.setup(opts.templatePath);
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
}
