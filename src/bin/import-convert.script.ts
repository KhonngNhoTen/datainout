#!/usr/bin/env node
import { program } from "commander";
import { ExcelImportTemplateGenerator } from "../importers/convert-file-import/ExcelImportTemplateGenerator.js";

program
  .description("Convert file excel to import description file")
  .option("-f, --file-sample <path>", "Specify the excel sample file")
  .option("-n, --name <path>", "Specify import template's path")
  .allowExcessArguments()
  .parse();

const opts = program.opts();
new ExcelImportTemplateGenerator(opts.name).write(opts.fileSample);
