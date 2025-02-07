import { CronJob, CronTime } from "cron";

export type HandlerCron = () => Promise<any>;
type EventCronName = "start" | "stop" | "tick" | "awake" | "error";
export class Cron {
  private scheduling: string | Date;
  name: string;
  private cronTime?: CronTime;
  private cronJob?: CronJob;
  private handlers: {
    [k in EventCronName]?: HandlerCron;
  } = {};

  constructor(scheduling: string | Date, name?: string, onTick?: HandlerCron, cronTime?: CronTime) {
    this.scheduling = scheduling;
    this.name = name ?? "";
    if (onTick) this.handlers.tick = onTick;
    this.cronTime = cronTime;
  }

  on(eventName: EventCronName, handler: HandlerCron) {
    if (eventName === "awake") this.onAwake(handler);
    if (eventName === "start") this.onStart(handler);
    if (eventName === "stop") this.onStop(handler);
    if (eventName === "tick") this.onTick(handler);
    if (eventName === "error") this.onError(handler);
  }

  async start() {
    if (!this.handlers.tick) throw new Error("onTick handler not set");
    this.cronJob = new CronJob(
      this.scheduling,
      this.handlers.tick,
      undefined,
      false,
      this.cronTime?.timeZone,
      null,
      !!this.handlers.awake,
      //   this.cronTime?.utcOffset,
      null,
      undefined,
      undefined,
      this.handlers.error
    );
    if (this.cronTime) this.cronJob.setTime(this.cronTime);

    if (this.handlers.awake) await this.handlers.awake();

    if (this.handlers.start) await this.handlers.start();
  }

  async stop() {
    if (this.handlers.stop) await this.handlers.stop();

    this.cronJob?.stop();
  }

  onStart(handler: HandlerCron) {
    this.handlers.start = handler;
  }

  onStop(handler: HandlerCron) {
    this.handlers.stop = handler;
  }

  onTick(handler: HandlerCron) {
    this.handlers.tick = handler;
  }

  onAwake(handler: HandlerCron) {
    this.handlers.awake = handler;
  }

  onError(handler: HandlerCron) {
    this.handlers.error = handler;
  }
}
