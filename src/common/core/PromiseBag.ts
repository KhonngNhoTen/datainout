export class PromiseBag {
  private promiseCount: number = 0;
  private promises: Promise<void>[] = [];
  constructor(promiseCount: number = 1) {}

  async addAndRun(callback: Promise<void>) {
    if (this.promiseCount === this.promises.length) {
      await Promise.all(this.promises);
      this.promises = [callback];
    } else this.promises.push(callback);
  }
}
