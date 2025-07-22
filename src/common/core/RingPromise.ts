import { Writable } from "stream";

type Task = (...args: any) => Promise<void>;

export class RingPromise {
  private ring: Promise<void>[] = [];
  private index = 0;
  private readonly task: Task;

  constructor(ringSize: number, task: Task) {
    this.ring = new Array(ringSize).fill(Promise.resolve());
    this.task = task;
  }

  async run(...arg: any): Promise<void> {
    this.index = (this.index + 1) % this.ring.length;
    await this.ring[this.index];
    this.ring[this.index] = this.task(...arg).catch((err) => {
      console.error(`Task at index ${this.index} failed:`, err);
    });
  }

  stream(): Writable {
    const that = this;

    return new Writable({
      objectMode: true,
      async write(arg, _encoding, callback) {
        try {
          await that.run(...arg);
        } catch (err) {
          return callback(err as any);
        }
        callback();
      },

      final(callback) {
        // Đợi tất cả các task còn lại trong vòng tròn kết thúc
        Promise.all(that.ring)
          .then(() => callback())
          .catch((err) => callback(err));
      },
    });
  }
}
