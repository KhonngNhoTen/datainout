import { BaseReader } from "../../importers/readers/BaseReader.js";
import { ReaderContainer } from "../../importers/readers/ReaderFactory.js";

export function IReader(name?: string, isDefault?: boolean) {
  return function (Class: any): any {
    return function () {
      const values = Object.getOwnPropertyDescriptors(Class);
      const nameClass = name ?? values.name.value;
      const reader = new (class extends Class {})();

      ReaderContainer.add(reader, nameClass, isDefault);
    };
  };
}
