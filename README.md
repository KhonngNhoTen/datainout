<div align="center">
  <a href="https://github.com/KhonngNhoTen/inoutjs">
    <img src="https://i.postimg.cc/P5WhbvfT/datainout-logo.png">
  </a>
  <h3>DataInOut</h3>
</div>  
<br>
This package enables data import and export (report-based) for Node.js applications. For importing, datainout uses Excel and CSV files to load data. For exporting, the package supports various file formats such as HTML, PDF, Excel, CSV, and more.

### Contents

1. [Quick start](#i-quicks-start)
2. [Configuration file](#ii-configuration-file)
3. [Excel file layout](#iii-excel-file-layout)
4. [Template generator](#iv-template-generator)
5. [Import](#v-import)
6. [Report](#vi-report)

### I. Quick start

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
    layoutDir: './excels',
  },
  report: {
    templateDir: './templates/reports',
    reportDir: './exports',
    layoutDir: './excels'
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
    layoutDir: './excels',
  },
  report: {
    templateDir: './templates/reports',
    reportDir: './exports',
    layoutDir: './excels'
  },
};
export default template;
```

In which:

- `templateExtension`: the type of file used.
- `templateDir`: the directory containing the generated template files (for import or report).
- `reportDir`: the location where exported files (HTML, PDF, Excel, etc.) are stored.
- `layoutDir`: The folder where the layout Excel (template) files are stored.

### III. Excel file layout

In datainout, Excel files are defined based on a fixed layout. Each Excel sheet is divided into three sections: header, table, and footer:

- `Header`: Defined as the rows from the beginning of the file up to (but not including) the title row of the main table.
- `Table`: The section that contains the main table data.
- `Footer`: The rows that come after the table content until the end of the sheet.

Symbols Used in the Excel Layout Definition File:

- `$`: Defines a variable in the header or footer section.
- `$$`: Defines a variable in the table section.
- `$$**`: Marks the start row of the table, and indicates that columns in this row are required (must have values) during import.

Syntax for Defining a Data Variable in a Cell:

```
<SYNTAX_TABLE>[field name]->[option1]&[option2]...
```

Example defining a table variable named title, with type string, and marked as required during import: `$$title->type=string&required`.

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

Then the layout file would look like this:

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

To create an Excel template file for both import and report purposes, you need to prepare a layout file that includes the header, footer, and table sections. Then, use **ExcelTemplateImport** (for import) or **ExcelTemplateReport** (for report) to generate the template file.

Below is the full code to generate the template:

```
import {
  ExcelTemplateImport,
  ExcelTemplateReport,
} from 'datainout/template-generators';

// Create import template file
async function createImportTemplate() {
  const sampleFilePath = './sampleFilePath.xlsx';
  const templatePath = './template-excel';
  await new ExcelTemplateImport(templatePath).generate(sampleFilePath);
}

// Create report template file
async function createReportTemplate() {
  const sampleFilePath = './sampleFilePath.xlsx';
  const templatePath = './template-excel';
  await new ExcelTemplateReport(templatePath).generate(sampleFilePath);
}
```

**Note:**
In your config file, you should set the paths to the layout file and template directories, so the paths used in code can be shorter and cleaner.

Or more simply, to generate a template file, you can use the following command:

```
npx datainout g [schema] -t [templateName] -s [sourceName]
```

**Where**:

- **schema:** `report` or `import`
- **templateName:** The name you want to give the generated template file (without extension).
- **sourceName:** The name of the layout file located in the layoutDir as defined in your config.

This command will read the layout from the layout file and generate a corresponding template (for import or report) into the templateDir folder specified in your config file.

Example:

```
npx datainout g -t user-import -s user-sample
```

This will generate a file named user-import.xlsx based on user-sample.xlsx.

### V. Import

```
type User = {
  name: string;
  password: string;
};
class Handler extends ImporterHandler<User> {
  protected async handleChunk(chunk: User[], filter: F) {
    // proccessing array of user
  }
}

async function import() {
  const templatePath = 'template-path';
  const filePath = 'file-path.xlsx';
  const importer = new Importer(templatePath);
  await importer.load(filePath, new Handler());
}
```

#### Handler:

The load function in the importer reads data from either a file path or a buffer, then passes chunks of data to a `handler` for processing.

The `handler` can be either:

- A **class** that extends the abstract class `ImporterHandler`, or

- A **list of functions** implementing the `ImporterHandlerFunction` signature.

This allows flexible handling of imported data in batches, making it easier to validate, transform, or insert into a database.

#### ImporterHandler:

##### Feature:

| **Key** | **Required** | **Note**                                                                  |
| ------- | ------------ | ------------------------------------------------------------------------- |
| eachRow | Optional     | **Boolean** – Indicates whether to process data row by row or in batches. |

##### Method:

**handleChunk(chunk: T[], filter: FilterImportHandler): Promise<void>**

Function to process data in chunks.

`Params:`

| Name   | Type                                        | Description                             |
| ------ | ------------------------------------------- | --------------------------------------- |
| chunk  | Array                                       | The chunk of data to be processed.      |
| filter | [FilterImportHandler](#filterimporthandler) | Detailed information of the data chunk. |

`Return:` Promise<void>

</br>

**handleRow(data: T, filter: FilterImportHandler): Promise<void>**

Function to process data row by row from the Excel file.

`Params:`

| Name   | Type                                        | Description                             |
| ------ | ------------------------------------------- | --------------------------------------- |
| chunk  | Array                                       | Data needs to be processed.             |
| filter | [FilterImportHandler](#filterimporthandler) | Detailed information of the data chunk. |

`Return:` Promise<void>

</br>

**handleHeader(header: any, filter: FilterImportHandler): Promise<void>**

Function to process header data.

`Params:`

| Name   | Type                                        | Description                             |
| ------ | ------------------------------------------- | --------------------------------------- |
| chunk  | Array                                       | The chunk of data to be processed.      |
| filter | [FilterImportHandler](#filterimporthandler) | Detailed information of the data chunk. |

`Return:` Promise<void>

</br>

**handleFooter(footer: any, filter: FilterImportHandler): Promise<void>**

Function to process footer data.

`Params:`

| Name   | Type                                        | Description                             |
| ------ | ------------------------------------------- | --------------------------------------- |
| chunk  | Array                                       | The chunk of data to be processed.      |
| filter | [FilterImportHandler](#filterimporthandler) | Detailed information of the data chunk. |

`Return:` Promise<void>

</br>

**catch(error: Error): Promise<void>**

Function to handle any errors that occur (including those thrown inside the `ImporterHandler` methods).

`Params:`

| Name  | Type  | Description            |
| ----- | ----- | ---------------------- |
| error | Error | Any errors that occur. |

`Return:` Promise<void>

#### List of ImporterHandlerFunction:

A handler can also be defined as an array of `ImporterHandlerFunction`. In this case, each data chunk (or error) will be passed through these functions sequentially. The output of one function will be used as the input for the next.

| Params | Type                                        | Note                                    |
| ------ | ------------------------------------------- | --------------------------------------- |
| data   | `TableData` or `Error` or `Error`[]         | data needs to process                   |
| filter | [FilterImportHandler](#filterimporthandler) | Detailed information of the data chunk. |

**Return:** `Any`.

##### FilterImportHandler

Detailed information for the filter:

- **sheetIndex:** `number` – The index of the sheet being processed.
- **sheetName?:** `string` or `undefined` – The name of the sheet being processed (optional).
- **section:** The section of the sheet currently being processed: `header`, `footer`, or `table`.
- **isHasNext:** `boolean` – Indicates whether there is more data to process or if this is the final chunk.

#### Class Importer:

**Method load():**

- arguments:

| Key  | Required | Note                                                      |
| ---- | -------- | --------------------------------------------------------- |
| arg1 | true     | **string** or **Buffer**                                  |
| arg2 | true     | **ImporterHandlerInstance**                               |
| arg3 | false    | [ImporterLoadFunctionOpions](#importerloadfunctionopions) |

- return: _void_

</br>

**Method createStream():**

- arguments:

| Key  | Required | Note                        |
| ---- | -------- | --------------------------- |
| arg1 | true     | **string** or **Readable**  |
| arg2 | true     | **ImporterHandlerInstance** |

- return: _IBaseStream_

##### ImporterLoadFunctionOpions

Details:

- type: "excel" or "csv". The type of file to import.
- chunkSize: Length of each chunk to import.
- ignoreErrors: Whether to ignore errors during import.
- jobCount: Number of concurrent jobs.

### VI. Report

To generate an export file, you can proceed as follows:

```
const templatePath = 'template-path';
const reportPath = 'report-path.xlsx';
const users = [
  // List of user
];
const exporter = new Reporter(templatePath).createExporterEXCEL();
await exporter.write(
  reportPath,
  { table: users }
);
```

#### Class Exporter:

**Method write():** Write the file to the specified path.

- arguments:

| Key  | Required | Note                                                                             |
| ---- | -------- | -------------------------------------------------------------------------------- |
| arg1 | true     | **string**                                                                       |
| arg2 | true     | [TableData](#tabledata) or [TablePartialDataTransfer](#tablepartialdatatransfer) |
| arg3 | false    | [ExcelExporterOptions](#excelexporteroptions)                                    |

- return: _void_

</br>

**Method toBuffer():** Generate the file and return it as a buffer.

- arguments:

| Key  | Required | Note                                                                             |
| ---- | -------- | -------------------------------------------------------------------------------- |
| arg1 | true     | [TableData](#tabledata) or [TablePartialDataTransfer](#tablepartialdatatransfer) |
| arg2 | false    | [ExcelExporterOptions](#excelexporteroptions)                                    |

- return: _Buffer_

</br>

**Method streamTo():** Create the file using a stream. Suitable for exporting large Excel datasets.

- arguments:

| Key  | Required | Note                                                  |
| ---- | -------- | ----------------------------------------------------- |
| arg1 | true     | **string** or **Writable**                            |
| arg2 | true     | [TablePartialDataTransfer](#tablepartialdatatransfer) |
| arg3 | false    | [ExcelExporterOptions](#excelexporteroptions)         |

- return: _void_

##### TableData

Details:

- header: Data of header sheet.
- footer: Data of footer sheet.
- table: Array of items.

##### TablePartialDataTransfer

Details:

- header: Data of header sheet.
- footer: Data of footer sheet.
- table: [PartialDataTransfer](#partialdatatransfer).

##### ExcelExporterOptions

Details:

- useSharedStrings: boolean.
- zip: Optional.
- style: "no-style" or "no-style-no-header" or "use-style".

#### PartialDataTransfer:

An abstract class responsible for retrieving data and piping it into an export file.

**Features:**

| Key          | Required | Note                                                            |
| ------------ | -------- | --------------------------------------------------------------- |
| **isStream** | false    | Boolean – Whether to fetch data using a stream.                 |
| **delayMs**  | false    | Number – Delay (in milliseconds) between each fetch.            |
| **jobCount** | false    | Number – Number of concurrent data-fetching jobs. Default is 1. |

**Methods:**

| Name                | Arguments | Note                                                 |
| ------------------- | --------- | ---------------------------------------------------- |
| **awake**           | none      | Function called before starting data retrieval.      |
| **completed**       | none      | Function called after completing data retrieval.     |
| **configSheetMeta** | string    | Configure mapping between the sheet and the dataset. |
| **fetchBatch**      | number    | Fetch a batch of data from the database.             |
| **createStream**    | none      | Create a data stream for export.                     |
