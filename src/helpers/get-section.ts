import { SheetSection } from "../common/types/common-type.js";
import * as exceljs from "exceljs";

export function getSection(rowIndex: number, beginTableAt: number, row: exceljs.Row, columnIndex: number): SheetSection {
  if (rowIndex < beginTableAt) return "header";
  if ((row.values as any[])[columnIndex]) return "table";
  return "footer";
}
