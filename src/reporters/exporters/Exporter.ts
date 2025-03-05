import { PassThrough } from "stream";
import { pathReport } from "../../helper/path-file.js";
import { CreateStreamOpts, ReportData } from "../type.js";
import { WriterStreanm } from "./stream/WriterStream.js";

export abstract class Exporter {
  protected templatePath: string;
  protected opts: any;
  constructor(templatePath: string, opts: any) {
    this.templatePath = pathReport(templatePath ?? "", "templateDir");

    this.opts = opts;
  }

  abstract writeFile(reportData: ReportData | ReportData[], path: string): Promise<any>;
  abstract buffer(reportData: ReportData | ReportData[]): Promise<Buffer>;

  abstract writerStream(opts: CreateStreamOpts): WriterStreanm;
}
