import { PassThrough } from "stream";

export abstract class ReportDataIterator {
  delay: number;
  constructor(delay: number) {
    this.delay = delay;
  }

  abstract next(): Promise<{ itemCount: number; data: any[]; hasNext: boolean }>;
  abstract reset(): void;

  createStream(): NodeJS.ReadableStream {
    const stream = new PassThrough();
    const pushData = async () => {
      let { data, hasNext } = await this.next();
      stream.push(Buffer.from(JSON.stringify(data)));
      if (!hasNext) {
        stream.push(null);
      } else setTimeout(pushData, this.delay);
    };

    pushData();
    return stream;
  }
}
