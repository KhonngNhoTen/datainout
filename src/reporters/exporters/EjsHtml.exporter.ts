import { PageData } from "../../common/types/common-type.js";
import { ExporterOptions } from "../../common/types/exporter.type.js";
import { Exporter } from "./Exporter.js";
import * as ejs from "ejs";

export class EjsHtmlExporter extends Exporter {
  constructor() {
    super(EjsHtmlExporter.name, "html");
  }
  async run(data: PageData, options: ExporterOptions): Promise<Buffer> {
    return Buffer.from(ejs.render(options?.templatePath ?? "", data));
  }
}
