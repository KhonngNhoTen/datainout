import { CellImportOptions } from "../common/types/import-template.type.js";
import { ValidateError } from "../common/error/ValidateError.js";
import { ERROR_MESSAGE } from "../common/core/error-message-default.js";

export function validateCellImport(value: any, cellDes: CellImportOptions, address: { col: string; row: number }) {
  const isNullValue = value === undefined || value === null;
  const validateErrorOpts = {
    col: address.col,
    row: address.row,
    value: value,
  };
  if (isNullValue && cellDes.required === true)
    throw new ValidateError({
      ...validateErrorOpts,
      message: `Cell [${address.col}${address.row}] ${ERROR_MESSAGE.MESSAGE_REQUIRED_CELL}`,
    });

  if (cellDes.validate) {
    const { isValid, message } = cellDes.validate(value);
    if (isValid)
      throw new ValidateError({
        ...validateErrorOpts,
        message: message ?? `Cell [${address.col}${address.row}]  ${ERROR_MESSAGE.MESSAGE_INVALID_CELL}`,
      });
  }
}
