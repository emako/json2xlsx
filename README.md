# JSONToXLSX

A useful method for editing an XLSX file instead of using local JSON files.

Supports conversion between JSON and XLSX, both ways.

This program is for use with vue-i18n.

## Usage

```bash
# Merge the JSON files into an XLSX file named output.xlsx.
json2xlsx en.json zh-Hans.json zh-Hant.json

# Extract XLSX to JSON files, with filenames based on column names.
json2xlsx output.xlsx
```

