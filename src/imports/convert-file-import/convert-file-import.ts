import * as exceljs from "exceljs";
import * as fs from "fs/promises";
import { CellDescription, CellType, ImportFileDesciptionOptions, SheetDesciptionOptions, SheetSection } from "../type";

/**
 * @param sourcePath file excel's path
 * @param outDir folder output indesc
 * @param name file'name output indesc
 */
export async function convertFileImport(sourcePath: string, outDir: string, name: string, beginTables: any[], endTables: any[]) {
  name = `${Date.now()}_${name}.indesc.js`;
  console.log(`Create imported file description: [${name}]`);
  const workBook = new exceljs.Workbook();
  await workBook.xlsx.readFile(sourcePath);

  const importDesciption: ImportFileDesciptionOptions = { sheets: [] };
  workBook.eachSheet((sheet, i) => {
    importDesciption.sheets.push(readSheet(sheet, beginTables[i - 1], endTables ? endTables[i - 1] : undefined));
  });

  await fs.writeFile(`${outDir}/${name}`, genContentFile(importDesciption), "utf-8");
  console.log(`Create file successfully!`);
}

function genContentFile(importDesciption: ImportFileDesciptionOptions): string {
  return `
  /** @type {import("inoutjs").ImportFileDesciptionOptions} */
  module.exports =
  ${JSON.stringify(importDesciption, null, undefined)}
  `;
}

function readSheet(sheet: exceljs.Worksheet, beginTable: number, endTable?: number) {
  const sheetDesciption: SheetDesciptionOptions = { content: [] };
  sheet.eachRow((row, rowIndex) => {
    row.eachCell((cell) => {
      const cellDescription = createCellDescription(cell, rowIndex, beginTable, endTable);
      if (cellDescription) sheetDesciption.content.push(cellDescription);
    });
    sheetDesciption.startTable = beginTable;
    sheetDesciption.endTable = endTable;
  });
  return sheetDesciption;
}

function getSection(rowIndex: number, beginTable: number, endTable?: number): SheetSection {
  let section: SheetSection = "table";
  if (rowIndex < beginTable) section = "header";
  else if (endTable && rowIndex > endTable) section = "footer";
  return section;
}

function getAddress(section: SheetSection, address: string) {
  return section === "table" ? address.split(/\d+/)[0] : address;
}

function createCellDescription(
  cell: exceljs.Cell,
  rowIndex: number,
  beginTable: number,
  endTable?: number,
): CellDescription | null {
  if (cell && (!(cell.value + "").includes("$") || (cell as any)._value.model.type === exceljs.ValueType.Merge)) return null;

  let cellValue = cell.value + "";
  let fieldName = "";
  let type: CellType = "string";
  const section: SheetSection = getSection(rowIndex, beginTable, endTable);

  if (cellValue.includes("$")) {
    fieldName = cellValue.split("$")[1];
    if (fieldName.includes("->")) {
      const args = fieldName.split("->");
      fieldName = args[0];
      type = args[1].toLowerCase() as CellType;
    }
  }

  return {
    address: getAddress(section, cell.address),
    section,
    fieldName,
    type,
  };
}
