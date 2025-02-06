import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellFormat, ExcelFormat, SheetFormat } from "../type";
import { SheetSection } from "../../type";
import { TemplateGenerator } from "./TemplateGenerator";

/**
 * Convert excel file into ExcelFormat type and save on template path
 */
export class Excel2ExcelTemplateGenerator extends TemplateGenerator {
  private beginTable: number | undefined;
  private endTable: number | undefined;
  private rowCount?: number;

  constructor(file: string, template: string) {
    super(file, template);
  }

  async generate() {
    console.log(this.file);
    const excelForm = await this.readWorkBook(this.file);

    const contentFile = `
     /** @type {import("inoutjs").ExcelFormat} */
  module.exports =
  ${JSON.stringify(excelForm, null, undefined)}
  `;
    await fs.writeFile(this.template, contentFile);
  }

  private async readWorkBook(file: string): Promise<ExcelFormat> {
    const workBook = new exceljs.Workbook();
    await workBook.xlsx.readFile(file);
    const excelForm: ExcelFormat = [];
    workBook.eachSheet((sheet) => {
      excelForm.push(this.readSheet(sheet));
    });

    return excelForm;
  }

  private readSheet(workSheet: exceljs.Worksheet) {
    this.beginTable = undefined;
    this.endTable = undefined;
    this.rowCount = workSheet.rowCount;
    const sheetForamt: SheetFormat = { cellFomats: [], beginTable: 1, rowHeights: {}, columnWidths: [] };
    workSheet.eachRow((row, rowIndex) => {
      this.setInforTable((row.values as any) ?? [], rowIndex);
      row.eachCell((cell) => {
        if ((cell as any)._value.model.type !== exceljs.ValueType.Merge && cell.value !== undefined) {
          const cellFormat = this.createCellFomat(rowIndex, cell);
          sheetForamt.cellFomats.push(cellFormat);
          if (cellFormat.section !== "table") {
            sheetForamt.rowHeights[rowIndex] = row.height;
          }
        }
      });
    });
    sheetForamt.columnWidths = workSheet.columns.map((col) => col.width);

    this.beginTable = this.beginTable ?? 1;
    sheetForamt.beginTable = this.beginTable;
    sheetForamt.endTable = this.endTable;

    return sheetForamt;
  }

  private setInforTable(values: any[], rowIndex: number) {
    for (let i = 0; i < values.length; i++) {
      const cell = values[i];
      const _isEditableCell = this.isEditableCell(cell);
      if (_isEditableCell) {
        if (!this.beginTable && (cell + "").includes("$$")) this.beginTable = rowIndex - 1;
        else if (this.beginTable && this.endTable === undefined && rowIndex - 1 !== this.beginTable && this.rowCount) {
          this.endTable = rowIndex - 1 - this.rowCount;
        }
      }
    }
  }

  private createCellFomat(rowIndex: number, cell: exceljs.Cell) {
    const cellValue = cell.value;
    const _isEditableCell = this.isEditableCell(cellValue);

    const section = this.getSection(rowIndex);

    const cellDesc: CellFormat = {
      address: section ? this.getAddress(cell.address, section, _isEditableCell) : cell.address,
      section,
      style: cell.style,
      value: this.getCellValue(cell),
      isHardCell: !_isEditableCell,
    };

    return cellDesc;
  }

  private getCellValue(cell: exceljs.Cell): CellFormat["value"] {
    if (!this.isEditableCell(cell.value)) return { hardValue: cell.value };
    const cellValue = (cell.value + "").replace("$$", "$");
    const args = cellValue.split("$")[1];
    let fieldName = args;
    if (args.includes("->")) fieldName = args.split("->")[0];

    return { fieldName };
  }

  private isEditableCell(cellValue: any) {
    return (cellValue + "").includes("$");
  }

  private getSection(rowIndex: number): SheetSection {
    if (!this.beginTable) return "header";

    let section: SheetSection = "table";
    if (rowIndex < this.beginTable) section = "header";
    else if (this.endTable && this.rowCount && rowIndex - 1 - this.rowCount >= this.endTable) section = "footer";
    return section;
  }

  private getAddress(address: string, section: SheetSection, isEditableCell: boolean) {
    return !isEditableCell ? address : section === "table" ? this.getColumnIndex(address) : address;
  }

  private getRowIndex(address: string) {
    return +address.split(/\w+/)[1];
  }

  private getColumnIndex(address: string) {
    return address.split(/\d+/)[0];
  }
}
