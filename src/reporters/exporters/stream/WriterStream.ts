import { PassThrough } from "stream";
import { CreateStreamOpts } from "../../type.js";

export interface WriterStreanm {
  add(chunks: any[], sheetIndex?: number): void;

  setContent(content: { sheetName?: string; header?: any; footer?: any }): void;

  /** data Finished */
  doneSheet(sheetIndex: number): Promise<void>;

  /** All data finished */
  allDone(): Promise<void>;

  stream(): PassThrough;
}
