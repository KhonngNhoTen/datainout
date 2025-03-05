import { program } from "commander";
import { Excel2ExcelTemplateGenerator } from "../../reporters/template-generator/ExcelTemplateGenerator.js";

program
  .description("Convert file excel into excel template file")
  .option("-f, --file <path>", "Specify the sample excel file")
  .option("-n, --name <char>", "Specify template file", "")
  .allowExcessArguments()
  .parse();
const opts = program.opts();
new Excel2ExcelTemplateGenerator(opts.file, opts.name).generate();
