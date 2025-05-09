import { Cron } from "./Cron.js";

export class CronManager {
  private crons: Record<string, Cron> = {};

  add(cron: Cron) {
    if (this.crons[cron.Key]) throw new Error(`Cron key ${cron.Key} is duplicated`);
    this.crons[cron.Key] = cron;
  }

  delete(key: string) {
    if (this.crons[key]) throw new Error(`Cron key ${key} is not found`);
    delete this.crons[key];
  }

  get(key: string) {
    if (this.crons[key]) throw new Error(`Cron key ${key} is not found`);
    return this.crons[key];
  }
}
