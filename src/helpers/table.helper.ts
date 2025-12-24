import { Cell, Row } from "exceljs";
import { AttributeType } from "../transformer/types/transformer-dto.js";
import { TABLE_KEY } from "../common/constant/constant.js";
import { TableScope } from "../template-mappers/types/template.type.js";

export class ExcelHelper {
    static getVariableValue(cellValue: any) {

        let fieldName = "";
        let type: AttributeType = "string";
        fieldName = cellValue.split("$")[1];
        if (fieldName.includes("->")) {
            const args = fieldName.split("->");
            fieldName = args[0];
            type = args[1].toLowerCase() as AttributeType;
        }

        return { fieldName, type };
    }

    static isVariableCell(cell: Cell) {
        const cellValue = cell.value + "";
        if (cellValue.includes(TABLE_KEY.INDEX_COLUMN_TABLE_SYNTAX)) return false;
        return cellValue.includes(TABLE_KEY.VARIABLE_SYNTAX);
    }

    static isVariableTable(cellValue: any) {
        return (cellValue + "").includes(TABLE_KEY.VARIABLE_TABLE_SYNTAX)
    }

    static getScope(rownumber: number, startAt: number, endAt?: number): TableScope {
        if(rownumber <= startAt) return "metadata";
        if(endAt && rownumber >= endAt) return "metadata";
        return "fields";
    }

}