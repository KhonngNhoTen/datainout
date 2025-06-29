import { CellImportOptions } from "../common/types/import-template.type.js";
import { ERROR_MESSAGE } from "../common/core/error-message-default.js";
import { ValidateImportError } from "../common/error/ValidateError.js";

export function validateCellImport(value: any, cellDes: CellImportOptions, address: string, keyField: string) {
  const isNullValue = value === undefined || value === null;
  const ValidateImportErrorOpts = {
    keyField,
    address,
    value: value,
  };
  if (isNullValue && cellDes.required === true) {
    const message = `${ERROR_MESSAGE.MESSAGE_REQUIRED_CELL}`;
    throw new ValidateImportError({
      ...ValidateImportErrorOpts,
      message,
    });
  }
  if (cellDes.validate) {
    const result = cellDes.validate(value);
    if (result instanceof Error) throw result;
    const { isValid, message } = result;
    if (isValid) {
      const defaultMessage = `${ERROR_MESSAGE.MESSAGE_INVALID_CELL}`;

      throw new ValidateImportError({
        ...ValidateImportErrorOpts,
        message: message ?? defaultMessage,
      });
    }
  }
}
