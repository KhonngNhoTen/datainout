import { ExporterFactory } from "./reporters/type.js";
export * from "./importers/type.js";

export * from "./reporters/type.js";

type ImportConfigOptions = {
  outDir?: string;
  templateDir?: string;
  sampleFileDir?: string;
};

type ReportConfigOptions = {
  templateDir?: string;
  reportDir?: string;
  excelSampleDir?: string;
  expoterFactory?: ExporterFactory;
};

export type ListOfPathImports = keyof Required<ImportConfigOptions>;
export type ListOfPathReports = keyof Omit<Required<ReportConfigOptions>, "expoterFactory">;

export type DataInoutConfigOptions = {
  import?: ImportConfigOptions;
  report?: ReportConfigOptions;
  dateFormat?: string;
  templateExtension?: ".ts" | ".js";
};

export type TableData = { header?: {}; footer?: {}; table?: any[] };
export type MultiTable = Array<TableData & { sheetName: string }>;
export type PageData = TableData | MultiTable | Record<string, any>;
