#!/usr/bin/env node
import * as path from "path";
import * as fs from "fs/promises";

const _path = path.join(process.cwd(), "datainout.config.js");

async function main() {
  fs.writeFile(
    _path,
    `/** @type {import("datainout").DataInoutConfigOptions} */
    module.exports = {
    dateFormat: "DD-MM-YYYY hh:mm:ss",
    templateExtension: ".js"
}`
  );
}

main();
