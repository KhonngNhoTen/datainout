export abstract class PartialDataTransfer {
  protected callBack?: (items: any[] | null, itemCount: number, total: number, hasNext: boolean) => Promise<any>;

  protected sleepTime: number = 200;

  constructor(sleepTime: number) {
    this.sleepTime = sleepTime;
  }

  start(callBack: (items: any[] | null, itemCount: number, total: number, hasNext: boolean) => Promise<any>) {
    this.callBack = callBack;
    (async () => {
      const { hasNext, itemCount, items, total } = await this.partialData();
      if (this.callBack) await this.callBack(items, itemCount, total, hasNext);
      await this.sleep(this.sleepTime);
    })();
  }

  sleep(ms: number) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  abstract partialData(): Promise<{ items: any[] | null; itemCount: number; total: number; hasNext: boolean }>;
}
