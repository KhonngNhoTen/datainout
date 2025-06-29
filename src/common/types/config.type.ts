type ImportConfigOptions = {
  /** Path of template files */
  templateDir?: string;

  /** Path of excel file to convert to template */
  excelSampleDir?: string;
};

type ReportConfigOptions = {
  /** Path of template files */
  templateDir?: string;

  /** Path of output file after generate report */
  reportDir?: string;

  /** Path of excel file to convert to template */
  excelSampleDir?: string;
  //   expoterFactory?: ExporterFactory;
};

export type ListOfPathImports = keyof Required<ImportConfigOptions>;
export type ListOfPathReports = keyof Omit<Required<ReportConfigOptions>, "expoterFactory">;

export type DataInoutConfigOptions = {
  import?: ImportConfigOptions;
  report?: ReportConfigOptions;
  dateFormat?: string;
  templateExtension?: ".ts" | ".js";
};
