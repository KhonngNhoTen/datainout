export class QueueData<T> {
  private stack: T[] = [];
  private size: number;
  private resolveFunc: () => void = () => {};
  private waitingFunc: Promise<void>;

  constructor(size = 50) {
    this.size = size;
    this.waitingFunc = this.createWaiter();
  }

  private createWaiter() {
    return new Promise<void>((resolve) => {
      this.resolveFunc = resolve;
    });
  }

  async waiting() {
    if (this.stack.length < this.size) return;
    await this.waitingFunc;
  }

  shift(): T | undefined {
    const data = this.stack.shift();
    if (this.stack.length === this.size - 1) {
      this.resolveFunc();
    }
    return data;
  }

  add(data: T) {
    this.stack.push(data);
    if (this.stack.length === this.size) {
      this.waitingFunc = this.createWaiter();
    }
  }
}
