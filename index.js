const { google } = require("googleapis");
const path = require("path");
const fs = require("fs");
var intToExcelCol = require("./intToExcelCol");
const keyfile = path.join(__dirname, "./oauth2.keys.json");
const tokenfile = path.join(__dirname, "./oauth2.token.json");

let sheet 

async function createSS(params) {
  var spreadsheetBody = {
    properties: {
      title: params.title || "New Sheet"
    },
    sheets: [
      {
        properties: {
          title: params.sheetTitle || "Pairs",
          gridProperties: {
            rowCount: params.rowCount || 1000,
            columnCount: 110
          }
        }
      }
    ]
  };

  return sheets.spreadsheets.create({ resource: spreadsheetBody }).then(
    function(response) {
      return response.data.spreadsheetId;
    },
    function(reason) {
      console.error("error: " + reason);
    }
  );
}

function updateSS(params) {
  return sheets.spreadsheets.values.update({
    spreadsheetId: params.spreadsheetId,
    range: intToExcelCol(params.startRow || 0) + (params.startRow + 1),
    valueInputOption: "USER_ENTERED",
    resource: {
      majorDimension: "ROWS",
      values: params.values
    }
  });
}

function batchUpdateSS(params, datumsToUpdate) {
  var batchPayload = {
    spreadsheetId: params.spreadsheetId,
    resource: {
      valueInputOption: "USER_ENTERED",
      data: datumsToUpdate.map(datum => ({
        range:
          (params.worksheetName ? `${params.worksheetName}!` : "") +
          datum.range,
        majorDimension: "ROWS",
        values: datum.values
      }))
      // majorDimension:
      // values: params.values
    }
  };
  console.log("batchPayload", JSON.stringify(batchPayload, null, 2));
  return sheets.spreadsheets.values.batchUpdate(batchPayload);
}

async function sendSpreadsheet(params) {
  var allRows = params.rows;

  var payloadRows = 36000;
  var spreadsheetId = params.spreadsheetId || null;

  for (var i = 0; i * payloadRows < allRows.length; i++) {
    var theseRows = allRows.slice(i * payloadRows, payloadRows * (i + 1));

    if (!spreadsheetId) {
      params.rows = theseRows;
      spreadsheetId = await createSS(params);
    }
    await updateSS({
      spreadsheetId,
      startRow: i * payloadRows + (params.startRow || 0),
      values: theseRows
    });
  }

  return spreadsheetId;
}

async function getSheetID(params) {
  var self = this;
  return new Promise((success, reject) => {
    sheets.spreadsheets.get(
      {
        spreadsheetId: params.spreadsheetId
      },
      (err, res) => {
        if (err) return console.log("The API returned an error: " + err);
        else {
          if (self && self.params && self.params.worksheetName) {
            success(
              res.data.sheets
                .filter(d => d.properties.title == self.params.worksheetName)
                .shift()
            );
          }

          success(res.data.sheets[0].properties.sheetId);
        }
      }
    );
  });
}

async function getWorksheetRange(params) {
  return new Promise((success, reject) => {
    sheets.spreadsheets.values.get(
      {
        spreadsheetId: params.spreadsheetId,
        range: params.range,
        valueRenderOption: params.valueRenderOption || "FORMATTED_VALUE"
      },
      (err, res) => {
        if (err) {
          reject({ spreadsheetId: params.spreadsheetId, err });
        } else {
          success(res.data.values);
        }
      }
    );
  });
}

async function getSpreadsheet(params) {
  const result = {};

  if (params.worksheetName) {
    params.range = `${params.worksheetName}!A1:ZZ5000`;
  } else {
    params.range = `A1:ZZ5000`;
  }

  params.valueRenderOption = "FORMATTED_VALUE";
  var arrayData = await getWorksheetRange(params);

  return arrayData;
}

async function saveRow() {
  var batch = [];
  console.log("<saveRow>", this.__row);
  if (this.__row == null) {
    await getRows.call(this.__s);
    this.__row = this.__s.__rowCount++;
  }

  Object.keys(this).forEach(attrib => {
    if (this.__s.__columnNames.indexOf(attrib) == -1) {
      this.__s.__columnNames.push(attrib);
      batch.push({
        range: intToExcelCol(this.__s.__columnNames.indexOf(attrib)) + "1",
        values: [[attrib]]
      });
      batch.push({
        range: intToExcelCol(this.__s.__columnNames.indexOf(attrib)) + "2",
        values: [[attrib]]
      });
    }

    if (this[attrib] != this.__original[attrib]) {
      batch.push({
        range:
          intToExcelCol(this.__s.__columnNames.indexOf(attrib)) +
          (this.__row + 3),
        values: [[this[attrib]]]
      });
    }
  });

  await batchUpdateSS(this.__s.__params, batch);
}

function Row(datum, row, s) {
  this.__original = {};
  this.__s = s;
  this.__row = row;

  this.save = saveRow;

  Object.keys(datum).forEach(column => {
    this[column] = datum[column];
    if (this.__row) {
      this.__original[column] = datum[column];
    }
  });

  var hide = ["save", "__row", "__original", "__s", "__columnNames"];

  hide.forEach(hideAttribute => {
    Object.defineProperty(this, hideAttribute, {
      enumerable: false,
      writable: true
    });
  });
}

function newRow(datum) {
  return new Row(datum, null, this);
}

async function getRows() {
  let arrayData = await getSpreadsheet(this.__params);

  this.__columnNames = arrayData.shift();

  arrayData.shift(); // skip human readable labels

  let objs = arrayData.map((data, i) => {
    return new Row(
      data.reduce((acc, datum, index) => {
        acc[this.__columnNames[index]] = datum;
        return acc;
      }, {}),
      i,
      this
    );
  });

  this.__rowCount = objs.filter(obj => Object.keys(obj).length > 0).length;

  return objs;
}

function Spreadsheet(params, creds) {
  if (!sheet) {
    let keys = creds.keys || JSON.parse(fs.readFileSync(keyfile));
    let token = creds.token || JSON.parse(fs.readFileSync(tokenfile));

    const oAuth2Client = new google.auth.OAuth2(
      keys.web.client_id,
      keys.web.client_secret
    );
    oAuth2Client.setCredentials(token);
    sheets = google.sheets({ version: "v4", auth: oAuth2Client });
  }

  var s = function(attribs) {
    if (attribs) {
      return newRow.call(s, attribs);
    } else {
      return getRows.call(s);
    }
  };

  s.__params = params;
  s.__columns = [];
  s.__rowsCount = 0;

  s.metadata = function() {
    return getSheetID.call(s, s.__params);
  };

  return s;
}

function formatDate(date) {
  return (
    parseInt(new Date(date).getDate()) +
    "/" +
    parseInt(new Date(date).getMonth() + 1) +
    "/" +
    parseInt(new Date(date).getFullYear())
  );
}

module.exports = {
  Spreadsheet
};

if (require.main == module) {
  (async () => {
    var params = {
      spreadsheetId: "1AwPH8WW3avCczXwhFYQ-AMhfSSABjO3W6pUmtqDBztk",
      worksheetName: "Users"
    };

    var sheet = Spreadsheet(params);

    // console.log("properties", await sheet.metadata());

    var rows = await sheet();
    console.log(rows);
    // rows[0].useCount++
    // rows[0].date = formatDate(new Date());

    // await rows[0].save();

    // var newRow = sheet({
    //   email: "three@test.com",
    //   useCount: 0
    // });
    // await newRow.save();

    // var newRow = sheet({
    //     email: "four@test.com",
    //     name: "Carlos",
    //     useCount: 0
    //   });
    //   await newRow.save();

    // var rows = await sheet();

    // var editRow = rows.filter((row) => row.email === 'two@test.com')[0]
    // editRow.useCount++;
    // editRow.date = '20/10/2019'

    // editRow.save()

    // var newRow = sheet({
    //     email: "yo@brother.com",
    //     lastName: 'steward',
    //     useCount: 19
    // });
    // await newRow.save();
  })();
}
