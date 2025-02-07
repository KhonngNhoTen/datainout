import { CronTime } from "cron";
import { Cron, HandlerCron } from "./Cron";

export class CronManager {
  private crons: Record<string, Cron> = {};
  async stopCron(name: string) {
    const cron = this.get(name);
    await cron.stop();
  }

  async startCron(name: string) {
    const cron = this.get(name);
    await cron.start();
  }

  createCron(scheduling: string | Date, name: string, onTick?: HandlerCron, cronTime?: CronTime) {
    if (this.crons.name) throw new Error(`Name cron[ ${name}] is dupplicated`);
    this.crons[name] = new Cron(scheduling, name, onTick, cronTime);

    return this.get(name);
  }

  get(name: string) {
    if (!this.crons.name) throw new Error(`Not found cron with name = ${name}`);
    return this.crons[name];
  }
}
