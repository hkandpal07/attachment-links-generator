# Attachments Link Generator

## Invocation 

```
Requirements: An excel file with at least the following two columns in the root of the program:
1. "Page ID"
2. "Attachment Title"
Column names are case sensitive
```

Invoke by passing file name and/or sheet name to command line args as follows

`node app.js <filename.xlsx> <sheetname.xlsx>`

Example usage: 

1. Will pick the first sheet in file if sheet name is not provided => `node app.js baseFile.xlsx`

2. Will pick data from the given sheet name if provided => `node app.js baseFile.xlsx Sheet1`

3. Enclose sheetName or fileName in double quotes if it has spaces => `node app.js "base File.xlsx" "Sheet Name"`

## Result

The resultant json will be printed in a file named `attachmentsData.json` in the root of the program.