
'use strict';

const {google} = require('googleapis');
const express = require('express');
const opn = require('open');
const path = require('path');
const fs = require('fs');

const keyfile = path.join(__dirname, './oauth2.keys.json');
const tokenfile = path.join(__dirname, './oauth2.token.json');
const keys = JSON.parse(fs.readFileSync(keyfile));
const scopes = ['https://www.googleapis.com/auth/spreadsheets'];

// Create an oAuth2 client to authorize the API call
const client = new google.auth.OAuth2(
  keys.web.client_id,
  keys.web.client_secret,
  'https://localhost:3000/oauth2callback'
);

// Generate the url that will be used for authorization
this.authorizeUrl = client.generateAuthUrl({
  access_type: 'offline',
  scope: scopes,
});

// Open an http server to accept the oauth callback. In this
// simple example, the only request to our webserver is to
// /oauth2callback?code=<code>
const app = express();
app.get('/oauth2callback', (req, res) => {
  const code = req.query.code;
  client.getToken(code, (err, tokens) => {
    if (err) {
      console.error('Error getting oAuth tokens:');
      throw err;
    }
    client.credentials = tokens;
    console.log("tokens", JSON.stringify(tokens, null, 2))
    res.send('Authentication successful! Please return to the console.');
    fs.writeFileSync(tokenfile, JSON.stringify(tokens, null, 2))
    console.log("Tokens saved in ", tokenfile)
    server.close();
    process.exit()
  });
});
const server = app.listen(3000, () => {
  // open the browser to the authorize url to start the workflow
  opn(this.authorizeUrl, {wait: false});
});
