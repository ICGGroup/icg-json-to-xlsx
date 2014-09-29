# icg-json-to-xlsx

A simple module to convert JSON to XLSX files.  This is really just a thin wrapper on the awesome [XLSX](https://github.com/SheetJS/js-xlsx) module.


## Installation

    [sudo] npm install icg-json-to-xlsx

## Overview

This was originally constructed for use in ICG's Email notification service to convert data pulled from an API into attachements that could be sent to users.  As such, it has been constructed with that primary use case in mind.

The icg-json-to-xlsx exports a function that accepts three paramaters

| Parameter   | Description                                                                                                                         |
|-------------|-------------------------------------------------------------------------------------------------------------------------------------|
| filename    | The path to the output file                                                                                                         |
| data        | The JSON data                                                                                                                       |
| options     | An options object that can contain a `headers` array with header values and a `sheetName` value for the title of the workbook sheet |
