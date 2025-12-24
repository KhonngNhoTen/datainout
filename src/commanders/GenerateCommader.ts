import { Command } from "commander";
import { AbstractCommander } from "./AbstractCommander.js";
import { GenerateAction } from "./command-actions/GenerateAction.js";

export class GenerateCommander extends AbstractCommander {
  constructor() {
    super("generate", new GenerateAction());
  }
  async run(program: Command) {
    program
      .command("generate <schema>")
      .alias("g")
      .description("Generate a new template")
      .option("--source-file", "Extension file source", "excel")
      .option("--out-file", "Extension file output", "excel")
      .option("-t, --name-template [nameTamplte]", "Path of file template", "")
      .option("-s, --name-source [nameSource]", "Path of source file", "")
      .allowExcessArguments()
      .action(async (...args) => await this.wrapAction(...args));
  }
}
