#!/usr/bin/env node
import { program } from "commander";
import { CommanderLoader } from "../commanders/CommanderLoader.js";
const bootstrap = async () => {
  program
    .version(require("../../package.json").version, "-v, --version", "Output the current version.")
    .usage("<command> [options]")
    .helpOption("-h, --help", "Output usage information.");

  await new CommanderLoader().load(program);
  await program.parseAsync();
};
bootstrap();
