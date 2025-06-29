import { Command } from "commander";
import { ICommandAction } from "./command-actions/ICommandAction.js";

export abstract class AbstractCommander {
  protected name: string;
  protected sleepTime: number = 3000;
  protected action: ICommandAction;
  protected logs = {
    Trigger: "",
    Run: "",
    Complated: "",
  };
  constructor(name: string, action: ICommandAction) {
    this.name = name;
    this.action = action;
    this.logs = {
      Complated: `Action ${this.name} is completed !!`,
      Run: `Action ${this.name} is running !!`,
      Trigger: `Action ${this.name} is waitting ${this.sleepTime}ms. \nTo cancel action, please press Ctrl + C...`,
    };
  }

  protected sleep() {
    return new Promise((resolve) => setTimeout(resolve, this.sleepTime));
  }

  protected async wrapAction(...data: any) {
    console.log(this.logs.Trigger);
    await this.sleep();
    console.log(this.logs.Run);
    await this.action.handleAction(...data);
    console.log(this.logs.Complated);
  }

  protected abstract run(program: Command): Promise<any>;
}
