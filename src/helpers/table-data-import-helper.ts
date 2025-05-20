import { SheetSection, TableData } from "../common/types/common-type.js";
import { BaseAttributeImporter, CellImportOptions, SheetImportOptions } from "../common/types/import-template.type.js";
import { CellDataHelper, RowDataHelper } from "./excel-reader-helper.js";
import { TypeParser } from "./parse-type.js";

/**
 * Class TableData mangements: push and get table.
 */
export class TableDataImportHelper {
  private container: { value: any; arrayValues: any[] } = { arrayValues: [], value: {} };
  private selectedSection?: SheetSection;
  private section?: SheetSection;
  private typeParser: TypeParser = new TypeParser();
  private endTableAt: number = -1;

  constructor(typeParser?: TypeParser) {
    this.typeParser = typeParser ?? new TypeParser();
  }

  /**
   * Push cells into TableData.
   *
   * Return true if the section is changed or the arrayValues length is greater than chunkSize.
   */
  push(
    cells: CellDataHelper[] | null,
    sheetImportOptions: SheetImportOptions,
    formattedCellImport: CellImportOptions[],
    chunkSize: number
  ): boolean {
    if (!cells) return true;
    const section = cells[0].section;
    if (!this.selectedSection) this.selectedSection = section;
    this.section = section;
    const trigger = this.isTriggerGroupValue(chunkSize, section, this.selectedSection);
    if (trigger && section === "footer") this.endTableAt = cells[0].rowIndex;

    const groupValues: any = {};

    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      const index = formattedCellImport.findIndex((e) => this.compareByAddress(e, cell, cell.rowIndex, this.endTableAt));
      if (index === -1) continue;
      const cellImport = formattedCellImport[index];
      groupValues[cellImport.keyName] = cell.detail.value;
    }

    if (trigger) {
      formattedCellImport.forEach((cell) => {
        const formatedValue = this.formatValue(cell, groupValues[cell.keyName], groupValues, this.typeParser);
        groupValues[cell.keyName] = formatedValue;
      });
    }
    if (section === "table" && Object.keys(groupValues).length > 0) this.container.arrayValues.push(groupValues);
    else if (Object.keys(groupValues).length > 0) this.container.value = { ...this.container.value, ...groupValues };

    return trigger;
  }

  /** get TableData */
  pop(): TableData {
    const result: TableData = { table: [], header: undefined, footer: undefined };

    if (this.selectedSection === "table") {
      result.table = this.container.arrayValues;
      this.container = { value: this.container.value, arrayValues: [] };
    } else if (this.selectedSection) {
      result[this.selectedSection] = this.container.value;
      this.container = { value: {}, arrayValues: this.container.arrayValues };
    }
    this.selectedSection = undefined;
    return result;
  }

  private formatValue(cell: BaseAttributeImporter, value: any, row: {}, typeParser: TypeParser) {
    if (cell.setValue) value = cell.setValue(value, row);
    if (cell.type && cell.type !== "virtual") value = (typeParser as any)[cell.type](value);
    if (cell.validate && cell.validate(value)) throw new Error("Validated fail");
    return value;
  }

  private compareByAddress(cell: CellImportOptions, cellRaw: CellDataHelper, rowNumber: number, endTableAt: number) {
    if (cell.section === "header") return cell.address === cellRaw.address;
    else if (cell.section === "table") return cellRaw.address === `${cell.address}${rowNumber}`;
    else {
      return rowNumber === endTableAt + (cell?.fullAddress?.row ?? 0);
    }
  }

  private isTriggerGroupValue(chunkSize: number, section?: SheetSection, selectedSection?: SheetSection) {
    return (
      this.selectedSection !== this.section || (this.container.arrayValues.length >= chunkSize && selectedSection === "table")
    );
  }
}
