import { RawRecordType } from "../../source/types/type.js";

export type AttributeType = "number" | "string" | "boolean" | "object" | "date" | "virtual";

// export type TableDto = {
//     fields: Record<string, any>,
//     metadata?: Record<string, any>;
// };
// export type EventDto = Record<string, any>;
export type MappedRecord = {
    type: RawRecordType;
    fields: Record<string, any>,
    metadata: Record<string, any>;
};