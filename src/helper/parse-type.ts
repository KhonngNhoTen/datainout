import dayjs from "dayjs";

type TypeParserOptions = {
  dateFormat?: string;
};
export class TypeParser {
  dateFormat: string;
  constructor(opts: TypeParserOptions) {
    this.dateFormat = opts?.dateFormat ?? "DD-MM-YYYY";
  }

  boolean(val: string) {
    return val === "true";
  }
  date(val: string) {
    return new Date(val); //dayjs(val, this.dateFormat).toDate();
  }
  number(val: string) {
    return +val;
  }
  object(val: string) {
    return JSON.parse(val);
  }
  string(val: string) {
    return val + "";
  }
}
