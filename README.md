<h1 align="center">oo-editors</h1>

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

Local browser-based document editor for Excel, Word, and PowerPoint files. Runs the [ONLYOFFICE](https://github.com/ONLYOFFICE/sdkjs) JavaScript SDK in a browser with a Node.js server handling format conversion.

Used as a downloadable extension in the Interpreter Desktop app.

## Document Flow

```
Open file
  → /open?filepath=/path/to/file.xlsx
  → x2t converts to ONLYOFFICE binary format
  → SDK renders document in browser (cell/word/slide editor)

Save file
  → SDK posts binary to /api/save
  → x2t converts back to XLSX/DOCX/PPTX
  → Written to original file path
```

## License

AGPL-3.0 -- see [LICENSE](LICENSE).

This project uses [ONLYOFFICE sdkjs](https://github.com/ONLYOFFICE/sdkjs) and [web-apps](https://github.com/ONLYOFFICE/web-apps), licensed under AGPL v3.
