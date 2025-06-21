import { BaseAttribute, SheetExcelOption, SheetSection } from "./common-type.js";

export type BaseAttributeImporter = {
  required?: boolean;
  setValue?: (attributeValue?: any, row?: Record<string, any>) => any;
  validate?: (val: any) => { isValid: boolean; message?: string };
  section: SheetSection;
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

export type SheetImportOptions = SheetExcelOption & { cells: CellImportOptions[] };

export type TableImportOptions = {
  sheets: SheetImportOptions[];
  name: string;
};
