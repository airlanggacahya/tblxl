
# tblxl
Library to convert dom table into excel xlsx files. Complete with styles and border conversion.

## Conversion
 * Border
 * Cell merge
 * Text aligment
 * Background color
 * Font color
 * Column width (in pixel)

## Usage
```js
var sheet = tblxl.table2sheet("table_id")
// you may want to set the column width pixel manually
tblxl.setColumnWidth(sheet, [80, 80, 120])
// trigger save as / download dialog
tblxl.save(sheet)
```

## Result
### Html
![html source](https://github.com/airlanggacahya/tblxl/blob/master/img/html.png)

### Excel Conversion
![excel conversion](https://github.com/airlanggacahya/tblxl/blob/master/img/excel.png)

## Dependency
tbl2xlsx depends on:
 * xlsx-js
 * Blob
 * FileSaver
 * jquery
 * jquery-cellpos
 * jszip (needed by xlsx-js)
 * lodash

All dependencies are located on /src/lib

## Run test page
Start static server on project root, look at [Big list of http static server](https://gist.github.com/willurd/5720255) for ideas. My favourite is python.
```bash
python -m http.server 8000
```

After that, open http://localhost:8000/test/index.html using some modern browser.
