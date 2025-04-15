import { PageData } from "../../common/types/common-type.js";
import { Exporter } from "./Exporter.js";
import * as ejs from "ejs";

export class EjsHtmlExporter extends Exporter {
  constructor() {
    super({ methodType: "full-load", name: EjsHtmlExporter.name, outputType: "html" });
  }
  async run(templatePath: string, data: PageData): Promise<Buffer> {
    return Buffer.from(ejs.render(templatePath, data));
  }
}
