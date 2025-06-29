<div align="center">
  <a href="https://github.com/KhonngNhoTen/inoutjs">
    <img src="https://i.postimg.cc/P5WhbvfT/datainout-logo.png">
  </a>
  <h3>DataInOut</h3>
</div>  
<br>
This package supports importing and exporting reports for Node.js. For importing, datainout uses Excel and CSV files to load data. For exporting, the package supports various file formats such as HTML, PDF, Excel, CSV, and more.

### Contents

1. [Quick start](#i-quicks-start)
2. [Configuration file](#ii-configuration-file)
3. [Excel file layout](#iii-excel-file-layout)
4. [Template generator](#iv-template-generator)

### I. Quicks start

- Install:

```
npm i datainout
```

- To generate the config file, please run the following command:

```
npx datainout init
```

### II. Configuration file

Javascript:

```
/** @type {import("datainout").DataInoutConfigOptions} */
module.exports = {
  templateExtension: '.js',
  import: {
    templateDir: './templates/imports',
    excelSampleDir: './excels',
  },
  report: {
    templateDir: './templates/reports',
    reportDir: './exports',
    excelSampleDir: './excels'
  },
};
```

TypeScript:

```
import { ExcelFormat } from "datainout";
const template : ExcelFormat = {
  templateExtension: '.ts',
  import: {
    templateDir: './templates/imports',
    excelSampleDir: './excels',
  },
  report: {
    templateDir: './templates/reports',
    reportDir: './exports',
    excelSampleDir: './excels'
  },
};
export default template;
```

In which:

- `templateExtension`: the type of file used.
- `templateDir`: the directory containing the generated template files (for import or report).
- `reportDir`: the location where exported files (HTML, PDF, Excel, etc.) are stored.
- `excelSampleDir`: the directory containing sample Excel files (template Excel files).

### III. Excel file layout

In datainout, Excel files are defined based on a fixed layout. Each Excel sheet is divided into three sections: header, table, and footer:

- `Header`: Defined as the rows from the beginning of the file up to (but not including) the title row of the main table.
- `Table`: The section that contains the main table data.
- `Footer`: The rows that come after the table content until the end of the sheet.

Symbols Used in the Sample Excel File (Excel Layout Definition File):

- `$`: Defines a variable in the header or footer section.
- `$$`: Defines a variable in the table section.
- `$$**`: Marks the start row of the table, and indicates that columns in this row are required (must have values) during import.

Syntax for Defining a Data Variable in a Cell:

```
<SYNTAX_TABLE>[field name]->[option1]&[option2]...
```

Example to defines a table variable named title, with type string and marked as required during import: `$$title->type=string&required`.

Available Options:

- type: Data type. Supported types include: string, number, date, boolean.
- required: Specifies that the field is required (only applies to import).
- default: Default value (only applies to import).

`Note`: You can use `$name->string` to define the variable cell, with name field is `name` and type is `string`.

The example:

|               |              |               |
| ------------- | ------------ | ------------- |
| Title of file |              |               |
| import date:  | 29-06-2025   |               |
|               |              |               |
| **Name**      | **Password** | **Email**     |
| Name 1        | 12345        | 123@email.com |
| Name 2        | 12345        | 124@email.com |
| Name 3        | 12345        | 125@email.com |
| -             | -            | -             |
|               | Created by   | Admin         |

If you have an Excel file as described, and the data you want to import includes:

- File title, Import date
- name, password, email
- Created by
- And the column "**name**" is always **required**.

Then the sample file (layout definition file) would look like this:

|                  |                     |                    |
| ---------------- | ------------------- | ------------------ |
| $title->string   |                     |                    |
| import date:     | $importDate->string |                    |
|                  |                     |                    |
| **$$\*\*->Name** | **Password**        | **Email**          |
| $$name->string   | $$password->string  | $$email->string    |
|                  |                     |                    |
|                  | Created by          | $createdBy->string |

Thus, the three layout sections of the above Excel file are:

- Header: **title**, **importDate**.
- Table: **name**, **password**, and **email**
- Footer: **createdBy**

### IV. Template generator
