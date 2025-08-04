type ModeSort = "DESC" | "ASC";

function reverseString(str: string) {
  return str.split("").reverse().join("");
}
export function sortByAddress<T extends { address?: string }>(arr: T[], mode: ModeSort = "ASC") {
  return arr.sort((a, b) => {
    if (!a.address) return 1;
    if (!b.address) return -1;
    if (a.address === b.address) return 0;
    const result = reverseString(a.address).localeCompare(reverseString(b.address));
    return result * (mode === "DESC" ? -1 : 1);
  });
}
