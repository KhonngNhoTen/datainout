import * as path from "path";
import { getConfig } from "../common/datainout.config.js";
import { DataInoutConfigOptions, ListOfPathImports, ListOfPathReports } from "../common/type.js";

const config = getConfig();

export class PathHelper {
    static getPath(mode: "import", _path: string, fieldName?: ListOfPathImports): string
    static getPath(mode: "report", _path: string, fieldName?: ListOfPathReports): string
    static getPath(mode: "import" | "report", _path: string, fieldName?: ListOfPathImports | ListOfPathReports): string {
        let path = ""
        if(mode === "import") path = this.pathImport(_path, fieldName as ListOfPathImports);
        else if(mode === "report") path = this.pathReport(_path, fieldName as ListOfPathReports);
        const ext =getConfig().templateExtension ?? "js";
        return `${path}${ext}`;
    }

    private static pathReport(_path: string, fieldName?: ListOfPathReports) {
        if (!config?.report || !fieldName || !config?.report?.[fieldName]) return path.join(process.cwd(), _path);
        return path.join(process.cwd(), config.report[fieldName], _path);
    }

    private static pathImport(_path: string, fieldName?: ListOfPathImports) {
        if (!config?.import || !fieldName || !config?.import?.[fieldName]) return path.join(process.cwd(), _path);
        return path.join(process.cwd(), config.import[fieldName], _path);
    }

}
