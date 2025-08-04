import { AttributeType, BaseAttribute, SheetExcelOption, SheetExcelOptionV2, SheetSection } from "./common-type.js";

export type BaseAttributeImporter = {
  type: AttributeType;
  required?: boolean;
  setValue?: (attributeValue?: any, row?: Record<string, any>) => any;
  validate?: (val: any) => { isValid: boolean; message?: string } | Error;
} & BaseAttribute;

/** Excel import template */
export type CellImportOptions = {
  address?: string;
  fullAddress?: {
    sheetName: string;
    address: string;
    row: number;
    col: number;
  };
} & BaseAttributeImporter;

export type SheetImportOptions = SheetExcelOption<CellImportOptions>;
