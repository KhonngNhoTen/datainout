import { ImporterReaderType } from "../../common/types/importer.type.js";
import { ReaderFactoryItem } from "../../common/types/reader.type.js";
import { BaseReader } from "./BaseReader.js";
import { ExcelJsCsvReader } from "./csv/ExceljsCsvReader.js";
import { ExcelJsReader } from "./exceljs/ExcelJsReader.js";
import { ExcelJsStreamReader } from "./exceljs/ExcelJsStreamReader.js";

export class ReaderFactory {
  protected readers: { [k in ImporterReaderType]: ReaderFactoryItem[] } = {
    excel: [],
    "excel-stream": [],
    csv: [],
  };

  constructor() {
    this.readers = {
      "excel-stream": [{ reader: new ExcelJsStreamReader(), isDefault: true, name: ExcelJsStreamReader.name }],
      csv: [{ reader: new ExcelJsCsvReader(), isDefault: true, name: ExcelJsCsvReader.name }],
      excel: [{ reader: new ExcelJsReader(), isDefault: true, name: ExcelJsReader.name }],
    };
  }

  add(reader: BaseReader, isDefault?: boolean, name?: string) {
    const type = reader.getType();
    isDefault = isDefault ?? false;
    if (!this.readers[type]) {
      this.readers[type] = [];
      isDefault = true;
    }
    if (isDefault) this.readers[type] = this.readers[type].map((e) => ({ ...e, isDefault: false }));

    this.readers[type].push({ isDefault, reader, name: name ?? reader.constructor.name });
  }

  get(type: ImporterReaderType) {
    const readers = this.readers[type];
    const index = readers.findIndex((e) => e.isDefault === true);
    if (index < 0) throw new Error(`Not found ImporterReader with type ${type}`);
    return this.readers[type][index];
  }

  getAll(type: ImporterReaderType) {
    return this.readers[type];
  }

  getByName(name: string) {
    for (const [type, readers] of Object.entries(this.readers)) {
      const index = readers.findIndex((e) => e.name === name);
      if (index >= 0) return readers[index];
    }
    throw new Error(`Not found ImporterReader with name ${name}`);
  }

  set(name: string, newValue: Partial<ReaderFactoryItem>) {
    const reader = this.getByName(name);
    if (newValue.name) reader.name = newValue.name;
    if (newValue.reader) reader.reader = newValue.reader;
    reader.isDefault = newValue.isDefault;
  }
}

export const ReaderContainer = new ReaderFactory();
