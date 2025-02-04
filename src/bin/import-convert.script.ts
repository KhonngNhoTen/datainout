import { program } from "commander";
import { convertFileImport } from "../imports/convert-file-import/convert-file-import";

program
  .description("Convert file excel to import description file")
  .option("-f, --file <path>", "Specify the excel template file")
  .option("-o, --outdir <path>", "Specify out folder")
  .option("-n, --name <char>", "File description name", "")
  .option("-b, --begin-row <items>", "Row's index begin table of each sheet. Example: 'X,X,X,...'", [])
  .option("-e, --end-row <items>", "Row's index end table of each sheet. Example: 'X,X,X,...'", [])
  .allowExcessArguments()
  .action((options) => {
    if (typeof options.beginRow === "string" && options.beginRow)
      options.beginRow = options.beginRow.split(",").map((val: string) => (val ? +val : undefined));
    if (typeof options.endRow === "string" && options.endRow)
      options.endRow = options.endRow.split(",").map((val: string) => (val ? +val : undefined));
  })
  .parse();

const opts = program.opts();
convertFileImport(opts.source, opts.outdir, opts.name, opts.beginRow, opts.endRow);
