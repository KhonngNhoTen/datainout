import { BaseReader } from "../../importers/readers/BaseReader.js";
import { ReaderContainer } from "../../importers/readers/ReaderFactory.js";

export function IReader(isDefault?: boolean, name?: string) {
  return function (Class: any): any {
    return function () {
      const values = Object.getOwnPropertyDescriptors(Class);
      const nameClass = name ?? values.name.value;
      const reader = new (class extends Class {})();

      ReaderContainer.add(reader, isDefault, nameClass);
    };
  };
}
