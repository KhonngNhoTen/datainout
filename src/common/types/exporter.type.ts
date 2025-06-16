export type ExporterOutputType = "csv" | "excel" | "html" | "pdf";
export type ExporterStreamOutputType = "excel";
export type ExporterMethodType = "full-load" | "stream";

export type ExporterOptions = {
  name: string;
  outputType: ExporterOutputType;
  methodType?: ExporterMethodType;
};
