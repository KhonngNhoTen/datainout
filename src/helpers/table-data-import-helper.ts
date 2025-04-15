import { SheetSection, TableData } from "../common/types/common-type.js";
import { BaseAttributeImporter, CellImportOptions, SheetImportOptions } from "../common/types/import-template.type.js";
import { CellDataHelper, RowDataHelper } from "./excel-reader-helper.js";
import { TypeParser } from "./parse-type.js";

export class TableDataImportHelper {
  private container: { value: any; arrayValues: any[] } = { arrayValues: [], value: {} };
  private selectedSection?: SheetSection;
  private section?: SheetSection;
  private typeParser: TypeParser = new TypeParser();

  constructor(typeParser?: TypeParser) {
    this.typeParser = typeParser ?? new TypeParser();
  }

  push(
    cells: CellDataHelper[] | null,
    sheetImportOptions: SheetImportOptions,
    formattedCellImport: CellImportOptions[],
    chunkSize: number
  ): boolean {
    if (!cells) return true;
    const section = cells[0].section;
    if (!this.section) this.selectedSection = section;
    this.section = section;

    const groupValues: any = {};

    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      const index = formattedCellImport.findIndex(
        (e) => () => this.compareByAddress(e, cell, cell.rowIndex, sheetImportOptions.endTableAt)
      );

      if (index === -1) continue;
      const cellImport = formattedCellImport[index];
      groupValues[cellImport.keyName] = cell.detail.value;
    }

    formattedCellImport.forEach((cell) => {
      const formatedValue = this.formatValue(cell, groupValues[cell.keyName], groupValues, this.typeParser);
      groupValues[cell.keyName] = formatedValue;
    });

    if (section === "table") this.container.arrayValues.push(groupValues);
    else this.container.value = { ...this.container.value, ...groupValues };

    return this.isTriggerGroupValue(chunkSize, section, this.section);
  }

  get(): TableData {
    const result: TableData = { table: [], header: undefined, footer: undefined };
    if (this.selectedSection === "table") {
      result.table = this.container.arrayValues;
      this.container = { value: this.container.value, arrayValues: [] };
    }
    if (this.selectedSection) {
      result[this.selectedSection] = this.container.value;
      this.container = { value: {}, arrayValues: this.container.arrayValues };
    }

    return result;
  }

  private formatValue(cell: BaseAttributeImporter, value: any, row: {}, typeParser: TypeParser) {
    if (cell.setValue) value = cell.setValue(value, row);
    if (cell.type && cell.type !== "virtual") value = (typeParser as any)[cell.type](value);
    if (cell.validate && cell.validate(value)) throw new Error("Validated fail");
    return value;
  }

  private compareByAddress(cell: CellImportOptions, cellRaw: CellDataHelper, rowNumber: number, endTableAt: number = -1) {
    if (cell.section === "header") return cell.address === cellRaw.address;
    else if (cell.section === "table") return cell.address === `${cellRaw.address}${rowNumber}`;
    else return cell.address === `${cellRaw.address}${endTableAt + rowNumber}`;
  }

  private isTriggerGroupValue(chunkSize: number, section?: SheetSection, selectedSection?: SheetSection) {
    return (
      this.selectedSection !== this.section || (this.container.arrayValues.length >= chunkSize && selectedSection === "table")
    );
  }
}
