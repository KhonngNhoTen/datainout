import { Command } from "commander";
import { AbstractCommander } from "./AbstractCommander.js";
import { InitializeAction } from "./command-actions/InitializeAction.js";

export class InitializeCommander extends AbstractCommander {
  constructor() {
    super("init", new InitializeAction());
  }
  async run(program: Command) {
    program
      .command("init")
      .alias("i")
      .option("-t, --type [type]", "Config file extension", "js")
      .description("Initialize the configuration file datainout")
      .allowExcessArguments()
      .action(async (...args) => await this.wrapAction(...args));
  }
}
