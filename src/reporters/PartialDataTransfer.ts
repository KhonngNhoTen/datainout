export abstract class PartialDataTransfer {
  private callBack?: (items: any[] | null, hasNext: boolean, isNewSheet: boolean) => Promise<any>;

  private sleepTime: number = 60;

  constructor(sleepTime: number) {
    this.sleepTime = sleepTime;
  }

  async start(callBack: (items: any[] | null, hasNext: boolean, isNewSheet: boolean) => Promise<any>) {
    this.callBack = callBack;
    let isLoop = true;
    while (isLoop) {
      const { items, hasNext } = await this.partialData();
      if (this.callBack) await this.callBack(items, hasNext, this.isNewSheet());
      await new Promise((resolve) => setTimeout(resolve, this.sleepTime));
      isLoop = hasNext;
    }
  }
  abstract partialData(): Promise<{ items: any[] | null; hasNext: boolean }>;

  protected isNewSheet(): boolean {
    return false;
  }
}
