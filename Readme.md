# Google Spreadsheet as DB (db-google-spreadsheet)

Use Google Spreadsheet as DB and Worksheets as Tables:

## Setup

Create `oauth2.keys.json` file with your _Google App_ *id* and *secret* as shown in `example_oauth2.keys.json`

You may need a CLIENT_ID, CLIENT_SECRET and REDIRECT_URL. You can find these pieces of information by going to the [Developer Console](https://console.developers.google.com/), clicking `your project --> APIs & auth --> credentials`.

Then execute `quickstart.js` file to get the `oauth2.token.json` file and use the module.

## How to use

### Open the Spreadsheet

```js
    var params = {
      spreadsheetId: "189ubij3PIK7ujsoSXXXXr3hsZqap_4w",
      worksheetName: "Users"
    };
    var sheet = Spreadsheet(params);
```

Example Spreadsheet Data:

![Sheet Data](https://user-images.githubusercontent.com/82532/67162710-8c6deb00-f35e-11e9-89f9-fea65c341be0.png "Sheet Data")

### Get Rows as Objects:

```js
var rows = await sheet();
console.log(rows)
```

_Console output:_
```sh
[ Row { email: 'one@test.com', useCount: '0' },
  Row { email: 'two@test.com', useCount: '0' },
  Row { email: 'three@test.com', useCount: '0' } ]
```

### Update Data:

```js
var editRow = rows.filter((row) => row.email === 'two@test.com')[0]
editRow.useCount++;
editRow.date = '20/10/2019'
editRow.save()
```

Updated Spreadsheet:

![Updated Spreadsheet](https://user-images.githubusercontent.com/82532/67162736-ebcbfb00-f35e-11e9-9643-0b7f1434c0a1.png "Updated Spreadsheet")

Note that it automatically added the new *Date* column

### Add new Row to Spreadsheet:

```js
var newRow = sheet({
    email: "yo@brother.com",
    useCount: 19
});
await newRow.save();
```

Updated Spreadsheet:

![Updated Spreadsheet](https://user-images.githubusercontent.com/82532/67162787-888e9880-f35f-11e9-9da4-e82864f38e2a.png "Updated Spreadsheet")

Note that it automatically added the new *Date* column

### Check Rows as Objects:

```js
var rows = await sheet();
console.log(rows)
```

_Console output:_
```sh
[ Row { email: 'one@test.com', useCount: '0' },
  Row { email: 'two@test.com', useCount: '1', date: '20/10/2019' },
  Row { email: 'three@test.com', useCount: '0' },
  Row { email: 'yo@brother.com', useCount: '19' } ]
```

## Column Attribute Mapping

The first row is reserved for the column names, so that order of columns and row two is used for human-friendly names that can be modified without affecting the DB.

![Attribute Mapping](https://user-images.githubusercontent.com/82532/67162823-e02d0400-f35f-11e9-8b1e-c87d1fd2a1f9.png  "Attribute Mapping")

This first row can be hidden for better visualization with `right-click -> hide Row`:

![Hide first Row](https://user-images.githubusercontent.com/82532/67162855-4154d780-f360-11e9-95a5-b9b906b85fc3.png  "Hide first Row")

