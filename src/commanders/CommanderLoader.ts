import { Command } from "commander";
import { InitializeCommander } from "./InitializeCommander.js";
import { GenerateCommander } from "./GenerateCommader.js";

export class CommanderLoader {
  async load(program: Command) {
    await new InitializeCommander().run(program);
    await new GenerateCommander().run(program);
  }
}
