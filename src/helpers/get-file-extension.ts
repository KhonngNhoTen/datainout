export function getFileExtension(path: string): "js" | "ts" {
  const fileName = path.replace("\\", "/").split("/").pop();
  if (!fileName) throw new Error(`Path [${path}] invalid`);
  const extension = fileName.split(".").pop();
  if (!extension) throw new Error(`Path [${path}] invalid`);
  if (extension === "js") return "js";
  else if (extension === "ts") return "ts";
  else throw new Error(`Path [${path}] invalid`);
}
