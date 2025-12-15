export type RawRecordType = "event" | "table" | "object" ;
export type SourceType = "excel" | "csv" | "log" ;//| "log" | "api";

export type RawRecord = {
    type: RawRecordType;
    fields: Record<string, any> | any[];
    metadata?: Record<string, any> | any[];
}


