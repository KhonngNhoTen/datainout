export abstract class PartialDataTransfer {
  protected callBack?: (items: any[] | null, isNewSheet: boolean) => Promise<any>;

  protected sleepTime: number = 60;

  constructor(sleepTime: number) {
    this.sleepTime = sleepTime;
  }

  async start(callBack: (items: any[] | null, isNewSheet: boolean) => Promise<any>) {
    this.callBack = callBack;
    const { items, hasNext } = await this.partialData();
    while (hasNext) {
      if (this.callBack) await this.callBack(items, this.isNewSheet());
      await new Promise((resolve) => setTimeout(resolve, this.sleepTime));
    }
  }

  // private sleep(ms: number) {
  //   return new Promise((resolve) => setTimeout(resolve, ms));
  // }

  abstract partialData(): Promise<{ items: any[] | null; hasNext: boolean }>;

  protected isNewSheet(): boolean {
    return false;
  }
}
