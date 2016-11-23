import google  from 'googleapis';




exports.SheetApi = class SheetApi {

    // This class exposes the Google Sheet API v4

    constructor(spreadsheetId = '',
                key = {},
                scopes = ['https://www.googleapis.com/auth/spreadsheets'])
    {

        // Auth data

        this._key = key;
        this._auth = {};
        this._scopes = scopes;

        // Sheet data

        this._cachedData = {spreadsheetId: spreadsheetId};
        this._worksheets = [];
        this._response = {}; //temp
        
        // Sheet api Object
        
        this._sheetApi = google.sheets('v4');
        
    };

    // authentication and authorization methods

    authenticate = () => {

        // Check if already authenticated

        if ( this._auth.hasOwnProperty('clientId_') ) {
            return 'Already authenticated';
        }

        // Create a JWT client

        if ( !this._auth.hasOwnProperty('key') ) {

            const key = this._key;

            let jwtClient = new google.auth.JWT(key.client_email, null, key.private_key, this._scopes, null);

            this._auth = jwtClient;

        }

    }

    auth = () => {

        // Create a JWT client if it doesn't exist already

        if ( !this._auth.hasOwnProperty('credentials') ) {

            const key = this._key;

            let jwtClient = new google.auth.JWT(key.client_email, null, key.private_key, this._scopes, null);

            this._auth = jwtClient;

        }

        // Check if already authorized

        if ( isAuthorized(this._auth) ) {
            return 'Already authorized';
        }



        // Return a promise with authorization attempt


        return new Promise((resolve, reject) => {

            this._auth.authorize(function (err, tokens) {

                if (err) {
                    console.log('error during authorization: ', err);
                    reject(err);
                }

                resolve(tokens);

            });

        });


    }

    // getters and setters

    saveMetaData = (metaData) => {

        console.log(metaData);

    }

    // Retrieve sheet data

    getSheetMetaData = async () => {

        await this.auth();

        return new Promise((resolve, reject) => {

            this._sheetApi.spreadsheets.get({
                auth: this._auth,
                spreadsheetId: this._cachedData.spreadsheetId,
            }, (err, resp) => {

                if (err) {
                    console.log('Data Error :', err)
                    reject(err);
                }

                this._response = resp;

                this._cachedData = parseData(resp);

                resolve(resp);

            });

        });

    }

    getSheetValues = async (range,
                            majorDimension = 'ROWS',
                            valueRenderOption = 'FORMATTED_VALUE') => {

        await this.auth();

        return new Promise((resolve, reject) => {

            this._sheetApi.spreadsheets.values.get({
                auth: this._auth,
                spreadsheetId: this._cachedData.spreadsheetId,
                range: range,
                majorDimension: majorDimension,
                valueRenderOption: valueRenderOption
            }, (err, resp) => {

                if (err) {
                    console.log('Data Error :', err)
                    reject(err);
                }

                resolve(resp);

            });

        });



    }


    getData = async (includeGridData = true) => {

        // Return a promise with the full spreadsheet

        await this.auth();



            return new Promise((resolve, reject) => {

                this._sheetApi.spreadsheets.get({
                    auth: this._auth,
                    spreadsheetId: this._cachedData.spreadsheetId,
                    includeGridData: includeGridData
                }, (err, resp) => {

                    if (err) {
                        console.log('Data Error :', err)
                        reject(err);
                    }

                    this._worksheets = resp;

                    this._cachedData = parseData(resp);



                    resolve(resp);

                });

            });


    }

    // Manipulate Sheet Data

    updateValues = async (ranges, values =['']) => {

        if (!Array.isArray(ranges)) {
            ranges = [ranges];
        }

        if (!Array.isArray(values)) {
            values = [values];
        }

        if (ranges.length !== values.length ) {
            throw new Error('Ranges and Values length should match');
        }

        await this.auth();

        return ranges.map((range, index) => {

            console.log('range is ',range);

            console.log('value is ',values[index]);

            return new Promise((resolve, reject) => {

                this._sheetApi.spreadsheets.values.update({
                    auth: this._auth,
                    spreadsheetId: this._cachedData.spreadsheetId,
                    range: range,
                    valueInputOption: 'USER_ENTERED',
                    resource: {range: range,
                        majorDimension: 'ROWS',
                        values: [[values[index]]]}
                } ,(err, resp) => {

                    if (err) {
                        console.log('Data Error :', err)
                        reject(err);
                    }

                    resolve(resp);

                });

            });

        });
    }

    batchUpdateValues = async (updates,
                         majorDimension= 'ROWS',
                         valueInputOption= 'USER_ENTERED') => {

        // It takes in input the starting cell (range) and an array of values
        // It updates the cells with the values and return a promise.
        // Check Sheets API documentation for the options

        if (!Array.isArray(updates)) {
            updates = [updates];
        }

        await this.auth();
        

        let request =  {
            valueInputOption: valueInputOption,
            data: []
        };


        updates.map((update) => {
            request.data.push({
                range: update.range,
                values: [update.values],
                majorDimension: majorDimension
            });
        });

            return new Promise((resolve, reject) => {

                this._sheetApi.spreadsheets.values.batchUpdate({
                    auth: this._auth,
                    spreadsheetId: this._cachedData.spreadsheetId,
                    resource: request
                } ,(err, resp) => {

                    if (err) {
                        console.log('Data Error :', err)
                        reject(err);
                    }

                    resolve(resp);

                });

        });

    }


    // Manipulate SpreadSheet MetaData (title, locale, etc)

    updateSheet = async (requests) => {

        await this.auth();

        return new Promise((resolve, reject) => {

            this._sheetApi.spreadsheets.batchUpdate({
                auth: this._auth,
                spreadsheetId: this._cachedData.spreadsheetId,
                resource: {requests: requests}
            } ,(err, resp) => {

                if (err) {
                    console.log('Data Error :', err)
                    reject(err);
                }

                resolve(resp);

            });

        });

        

    }


}

parseData = (resp) => {

    let sheets = resp.sheets.map((sheet) => {

        return {
            workSheetId: sheet.properties.sheetId,
            title: sheet.properties.title,
            rowCount: sheet.properties.gridProperties.rowCount,
            colCount: sheet.properties.gridProperties.columnCount
        };


    });

    let data = [{}];

    if (resp.sheets[0].hasOwnProperty('data')) {

        // create an array with the parsed elements

        data = resp.sheets.map((sheet) => {
            return sheet.data[0].rowData.map((row, rowIndex) => {
                if (Array.isArray(row.values)) {
                    return row.values.map((value, colIndex) => {
                        if (value.hasOwnProperty('formattedValue')) {
                            return {value: value.formattedValue,
                            row: rowIndex,
                            col: colIndex};
                        }
                    });
                }
            });
        });

        // remove undefined rows

        data = data.map((sheet) => {
            return sheet.filter((rows) => {
                    return (rows !== undefined);
                });
            });

        // remove undefined items

        data = data.map((sheet) => {
            return sheet.map((rows) => {
                return rows.filter((elem) => {
                    return elem !== undefined;
                });
            });
        });

        // remove empty rows

        data = data.map((sheet) => {
            return sheet.filter((rows) => {
                return rows.length !== 0;
            });
        });

        // flatten data array

        data = data.map((sheet) => {
            return [].concat(...sheet);
        });


    }

    return {
        spreadsheetId: resp.spreadsheetId,
        title: resp.properties.title,
        worksheets: sheets,
        data: data
    };

}



isAuthorized = (JWT) => {
    return Object.prototype.hasOwnProperty.call(JWT.credentials ,'access_token');
}







