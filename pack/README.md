[![npm](https://img.shields.io/npm/v/json2xlsx.cli.svg)](https://www.npmjs.com/package/json2xlsx.cli) [![Platform](https://img.shields.io/badge/platform-Windows-blue?logo=windowsxp&color=1E9BFA)](https://dotnet.microsoft.com/zh-cn/download/dotnet/latest/runtime) [![Platform](https://img.shields.io/badge/platform-Linux-green?logo=linux&color=1E9BFA)](https://dotnet.microsoft.com/zh-cn/download/dotnet/latest/runtime) [![Platform](https://img.shields.io/badge/platform-macOS-lightgrey?logo=apple&color=1E9BFA)](https://dotnet.microsoft.com/zh-cn/download/dotnet/latest/runtime)

# JSONToXLSX

Editing vue-i18n locale files in XLSX, a  useful program for editing an XLSX file instead of using local JSON files.

Supports conversion between JSON and XLSX, both ways.

This program is for use with vue-i18n.

## Dependencies

| Target Runtime | Dependencies                                                 |
| -------------- | ------------------------------------------------------------ |
| win7-x86       | [.NET Framework 4.7.2](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net472) |
| others         | [.NET 9.0](https://dotnet.microsoft.com/en-us/download/dotnet/9.0) |

## Usage

Only supports the first-level keys of the JSON, does not support nested keys.

```bash
# Merge the JSON files into an XLSX file named output.xlsx.
json2xlsx en.json zh-Hans.json zh-Hant.json

# Extract XLSX to JSON files, with filenames based on column names.
json2xlsx output.xlsx
```

