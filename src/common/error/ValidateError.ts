import { ValidateErrorOpts } from "../types/error.type.js";

export class ValidateError extends Error {
  col: string;
  row: number;
  value: any;
  constructor(opts: ValidateErrorOpts) {
    super(opts.message);
    this.col = opts.col;
    this.row = opts.row;
    this.value = opts.value;
  }
}
