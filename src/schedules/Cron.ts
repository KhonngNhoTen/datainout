import { CronJob } from "cron";
import { CronOptions } from "../common/types/cron.type.js";

export class Cron {
  private cronJob: CronJob;
  private key: string;

  constructor(cronTime: string, key: string, onTick: (...args: any[]) => Promise<void>, options?: CronOptions) {
    this.cronJob = new CronJob(cronTime, onTick, null, null, null, null, options?.runOnInit);
    this.cronJob.runOnce = options?.runOnce ?? false;
    this.key = key;
  }

  start() {
    this.cronJob.start();
  }

  stop() {
    this.cronJob.stop();
  }

  public get Key(): string {
    return this.key;
  }
}
