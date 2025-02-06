import * as path from "path";
import { getConfig, InoutConfigOptions } from "../inoutjs-config";

const config = getConfig();
export function pathReport(_path: string, fieldName?: keyof Required<Required<InoutConfigOptions>["report"]>) {
  if (!config?.report || !fieldName || !config?.report?.[fieldName]) return path.join(process.cwd(), _path);
  return `${config.report[fieldName]}/${_path}`;
}
