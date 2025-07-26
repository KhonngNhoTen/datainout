import { Writable } from "stream";
import { Task } from "../common/types/common-type.js";
import { ISheetMeta, SheetMeta } from "./SheetMeta.js";
import { QueueData } from "./QueueData.js";

export interface IPartialDataHandler {
  do(args: any): Promise<boolean>;

  stream(): Writable;
}

export type DataTranfer = {
  items: any[] | null;
  sheetCompleted: boolean;
  isCompleted: boolean;
  sheetName?: string;
  jobIndex?: number;
};

export class PartialDataHandler implements IPartialDataHandler {
  private task: Task<DataTranfer>;
  private originalSheetName: string;
  private sheetMeta?: ISheetMeta;

  done: Task<any> = async () => {};

  set SheetMeta(value: ISheetMeta) {
    this.sheetMeta = value;
  }

  constructor(originalSheetName: string, task: Task<DataTranfer>) {
    this.task = task;
    this.originalSheetName = originalSheetName;
  }

  async do(args: Pick<DataTranfer, "items" | "jobIndex">): Promise<boolean> {
    let sheetName = this.originalSheetName;
    let isCompleted = args.items === null || args.items.length === 0;
    let sheetCompleted = isCompleted;
    if (this.sheetMeta) {
      this.sheetMeta.updateRowCount(isCompleted, args.items?.length);
      sheetName = this.sheetMeta.getSheetName(args.jobIndex);
      isCompleted = isCompleted ? isCompleted : this.sheetMeta.IsCompleted;
      sheetCompleted = sheetCompleted ? sheetCompleted : this.sheetMeta.getSheetStatus(sheetName);
      if (args.jobIndex && (args.items === null || args.items.length === 0)) this.sheetMeta.completeJob(args.jobIndex);
    }
    await this.task({ items: args.items, jobIndex: args.jobIndex, isCompleted, sheetCompleted, sheetName });
    if (isCompleted) await this.done(null);

    return isCompleted;
  }

  stream(): Writable {
    const that = this;

    return new Writable({
      objectMode: true,
      async write(arg, _encoding, callback) {
        try {
          await that.do({
            items: arg,
          });
        } catch (err) {
          return callback(err as any);
        }
        callback();
      },
    });
  }
}
