class SheetHeadedData {
    //sheetName: string;
    range: SSHeadedRange;

    values: [{ [index: string]: any }]
    headers: Array<string>;

    constructor(sheetName: string, range: SSHeadedRange) {
        this.sheetName = sheetName;
        this.range = range;
    }

    // HELPERS
    private _sheetName: string;
    private sheet: GoogleAppsScript.Spreadsheet.Sheet;

    get sheetName(): string {
        return this._sheetName;
    }

    set sheetName(newName: string) {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var newSheet = ss.getSheetByName(newName);
        if (newSheet == null) {
            throw new Error("Sheet name: " + newName + " does not exist");
        }

        this._sheetName = newName;
        this.sheet = newSheet;
    }

    // READING

    prepareRead() {
        var sheet = this.sheet;
        // Run some checks
        if (this.range.row == 0) {
            this.range.row = 1;
        }
        if (this.range.column == 0) {
            this.range.column = 1;
        }
        if (this.range.numRows == 0) {
            this.range.numRows = sheet.getMaxRows();
        }
        if (this.range.numColumns == 0) {
            this.range.numColumns = sheet.getMaxColumns();
        }
        if ((this.range.row + this.range.numRows) > sheet.getMaxRows()) {
            this.range.numRows = sheet.getMaxRows() - this.range.row + 1;
            Logger.log("Num rows out of bounds. Using Max Rows " + this.range.numRows + " of spreadsheet")
        }
        if ((this.range.column + this.range.numColumns) > sheet.getMaxColumns()) {
            this.range.numColumns = sheet.getMaxColumns() - this.range.column + 1;
            Logger.log("Num columns out of bounds. Using Max Columns " + this.range.numColumns + " of spreadsheet")
        }
        if (this.range.headerRow == 0) {
            this.range.headerRow = 1;
        }
        if (this.range.firstDataRow < 2) {
            Logger.log("Invalid First Data Row " + this.range.firstDataRow + ". Setting it to 2")
            this.range.firstDataRow = 2;
        }
        if (this.range.firstDataRow <= this.range.headerRow) {
            Logger.log("Invalid First Data Row " + this.range.firstDataRow + ". Has to be higher than headerRow " + this.range.headerRow + ". Setting it to " + this.range.headerRow + 1)
            this.range.firstDataRow = this.range.headerRow + 1;
        }
    }

    getHeaders() {
        this.prepareRead();
        var headerValues = this.headerRange().getValues()[0];
        var headers: Array<string> = [];
        for (var i = 0; i < headerValues.length; i++) {
            var header = String(headerValues[i]);
            if (header == "") {
                header = "#" + i;
            }
            headers[i] = header;
        }
        this.headers = headers;
    }

    getValues() {
        this.prepareRead();
        this.getHeaders();

        var dataValues = this.dataRange().getValues();
        var values: [{ [index: string]: any }] = [{}];
        values.pop();   // I have no other way to init this array and spend too much time messing with this bs
        var i = 0;

        dataValues.forEach(row => {
            values[i] = {};
            for (var columnNum = 0; columnNum < this.range.numColumns; columnNum++) {
                var header: string = this.headers[columnNum];
                values[i][header] = row[columnNum];
            }
            i++;
        });
        this.values = values;
    }

    // WRITING
    prepareWrite() {
        var sheet = this.sheet;
        // Run some checks
        if (this.range.row == 0) {
            this.range.row = 1;
        }
        if (this.range.column == 0) {
            this.range.column = 1;
        }
        if (this.range.numRows == 0) {
            this.range.numRows = sheet.getMaxRows();
        }
        if (this.range.numColumns == 0) {
            this.range.numColumns = sheet.getMaxColumns();
        }
        /* Don't need the following on write as the sheet will automatically extend.
        if ((this.range.row + this.range.numRows) > sheet.getMaxRows()) {
            this.range.numRows = sheet.getMaxRows() - this.range.row + 1;
            Logger.log("Num rows out of bounds. Using Max Rows " + this.range.numRows + " of spreadsheet")
        }
        if ((this.range.column + this.range.numColumns) > sheet.getMaxColumns()) {
            this.range.numColumns = sheet.getMaxColumns() - this.range.column + 1;
            Logger.log("Num columns out of bounds. Using Max Columns " + this.range.numColumns + " of spreadsheet")
        }*/
        if (this.range.headerRow == 0) {
            this.range.headerRow = 1;
        }
        if (this.range.firstDataRow < 2) {
            Logger.log("Invalid First Data Row " + this.range.firstDataRow + ". Setting it to 2")
            this.range.firstDataRow = 2;
        }
        if (this.range.firstDataRow <= this.range.headerRow) {
            Logger.log("Invalid First Data Row " + this.range.firstDataRow + ". Has to be higher than headerRow " + this.range.headerRow + ". Setting it to " + this.range.headerRow + 1)
            this.range.firstDataRow = this.range.headerRow + 1;
        }
    }

    headerRange(): GoogleAppsScript.Spreadsheet.Range {
        var sheet = this.sheet;
        var range = this.range;
        var headerRange = sheet.getRange(range.row + range.headerRow - 1, range.column, 1, range.numColumns);
        return headerRange;
    }

    writeHeaderRange(): GoogleAppsScript.Spreadsheet.Range {
        var sheet = this.sheet;
        var range = this.range;
        var headerRange = sheet.getRange(range.row + range.headerRow - 1, range.column, 1, this.headers.length);
        return headerRange;
    }

    dataRange(): GoogleAppsScript.Spreadsheet.Range {
        var sheet = this.sheet;
        var range = this.range;
        var dataRange = sheet.getRange(range.row + range.firstDataRow - 1, range.column, range.numRows - range.firstDataRow + 1, range.numColumns);
        return dataRange;
    }

    writeDataRange(): GoogleAppsScript.Spreadsheet.Range {
        var sheet = this.sheet;
        var range = this.range;
        var dataRange = sheet.getRange(range.row + range.firstDataRow - 1, range.column, this.values.length, this.headers.length);
        return dataRange;
    }


    writeHeaders() {
        // Write the headers
        this.prepareWrite();
        var headerArray = [this.headers];
        this.writeHeaderRange().setValues(headerArray);
    }
    clearValues() {
        // TODO: Implement
        //this.writeDataRange().clearContent(); Does not work

    }
    /**
     * @param  {Boolean=false} writeHeaders - false will use the current headers. true will create new or overwrite them
     */
    writeValues(writeHeaders: Boolean = false) {
        this.prepareWrite();
        if (writeHeaders) {
            this.writeHeaders();
        } else {
            this.getHeaders();
        }
        // Write Data

        var data = [];
        this.values.forEach(dataDictionary => {
            var dataArray = [];
            this.headers.forEach(key => {
                var value = dataDictionary[key] || "";
                dataArray.push(value);
            });
            data.push(dataArray);
        });
        this.writeDataRange().setValues(data);
    }

    appendValues(newValues: [{ [index: string]: any }]) {
        this.getValues();
        var lastRow = this.sheet.getLastRow() - 1;
        var row = lastRow;
        newValues.forEach(element => {
            this.values[row] = element;
            row++;
        });
    }

}

class SSRange {
    row: number;
    column: number;
    numRows: number;
    numColumns: number;

    constructor(row: number, column: number, numRows: number, numColumns: number) {
        this.row = row;
        this.column = column;
        this.numRows = numRows;
        this.numColumns = numColumns;
    }

}

class SSHeadedRange extends SSRange {

    constructor(row: number, column: number, numRows: number, numColumns: number, headerRow: number, firstDataRow: number) {
        super(row, column, numRows, numColumns);
        this.headerRow = headerRow;
        this.firstDataRow = firstDataRow;
    }

    headerRow: number;
    firstDataRow: number;
}

interface String {
    equals(compare: string): boolean;
  }
  
  String.prototype.equals = function (compare: string) : boolean {
    var s: string = String(this);

    if (s.match(compare) !== null) {
        return true;
    }
    return false;
  }