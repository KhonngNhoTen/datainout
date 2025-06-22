import { ValidateImportErrorOpts } from "../types/error.type.js";

export class ValidateImportError extends Error {
  address?: string;
  value?: any;
  keyField: string;
  constructor(opts: ValidateImportErrorOpts) {
    super(opts.message);
    this.address = opts?.address;
    this.keyField = opts.keyField;
    this.value = opts.value;
  }
}
