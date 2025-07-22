import { Readable } from "stream";
import { PartialDataHandler } from "./IPartialDataHandler.js";
import { SheetMeta, SheetMetaOptions } from "./SheetMeta.js";
export interface PartialDataTransferRunner {
  start(): Promise<void>;
  completed(): Promise<void>;
  init(partialDataHandler: PartialDataHandler, sheetName: string): Promise<void>;
}
export abstract class PartialDataTransfer {
  private delayMs: number = 5;
  private isStream: boolean = false;
  private partialDataHandler: PartialDataHandler = {} as any;
  private jobCount: number = 1;

  constructor(opts?: { isStream?: boolean; delayMs?: number; jobCount?: number }) {
    this.delayMs = opts?.delayMs ?? 5;
    this.isStream = opts?.isStream ?? false;
    this.jobCount = opts?.jobCount ?? 1;
  }

  private async init(partialDataHandler: PartialDataHandler, originalSheetName: string) {
    const sheetMetaOptions = this.configSheetMeta(originalSheetName);
    this.partialDataHandler = partialDataHandler;
    this.partialDataHandler.done = this.completed;
    if (sheetMetaOptions) this.partialDataHandler.SheetMeta = new SheetMeta(sheetMetaOptions) as any;
    await this.awake();
  }

  private async start() {
    if (this.isStream) await this.startStream();
    else await this.startJobs();
  }

  /** Run with stream */
  private async startStream() {
    const readable = this.createStream();
    if (readable === null) throw new Error("You must implement 'createStream' method when using isStream = true");
    const wriable = this.partialDataHandler.stream();
    readable.pipe(wriable);
  }

  /** Run with one or multiples job */
  private async startJobs() {
    const promises = [];
    for (let i = 0; i < this.jobCount; i++) promises.push(async () => await this.createJob(i));
    await Promise.all(promises);
  }

  private async createJob(i: number) {
    let isLoop = true;
    while (isLoop) {
      const { hasNext, items } = await this.fetchBatch();
      this.partialDataHandler.do({ items, jobIndex: i });
      isLoop = hasNext;
      await new Promise((resolve) => setTimeout(resolve, this.delayMs));
    }
  }

  protected configSheetMeta(originalSheetName: string): SheetMetaOptions[] {
    return undefined as any;
  }
  async awake() {}
  async completed() {}

  /** Batching data */
  protected async fetchBatch(): Promise<{ items: any[] | null; hasNext: boolean }> {
    return { items: null, hasNext: false };
  }

  /** Run with streaming data */
  protected createStream(): Readable {
    return null as any;
  }
}
