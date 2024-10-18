# JSONToXLSX

A useful program for editing an XLSX file instead of using local JSON files.

Supports conversion between JSON and XLSX, both ways.

This program is for use with vue-i18n.

## Dependencies

| Target Runtime | Dependencies                                                 |
| -------------- | ------------------------------------------------------------ |
| win7-x86       | [.NET Framework 4.7.2](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net472) |
| others         | [.NET 8.0](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) |

## Usage

Only supports the first-level keys of the JSON, does not support nested keys.

```bash
# Merge the JSON files into an XLSX file named output.xlsx.
json2xlsx en.json zh-Hans.json zh-Hant.json

# Extract XLSX to JSON files, with filenames based on column names.
json2xlsx output.xlsx
```

