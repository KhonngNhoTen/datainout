import { ImporterHandler } from "../../importers/ImportHandler.js";
import { BaseAttribute, SheetSection } from "./common-type.js";
import { ImporterBaseReaderType } from "./importer.type.js";

export type BaseAttributeImporter = {
  setValue?: (attributeValue?: any, row?: Record<string, any>) => any;
  validate?: (val: any) => boolean;
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

export type SheetImportOptions = {
  cells: CellImportOptions[];
  sheetIndex: number;
  sheetName: string;
  beginTableAt: number;
  endTableAt?: number;
  keyTableAt: number;
};

export type TableImportOptions = {
  sheets: SheetImportOptions[];
  name: string;
};
