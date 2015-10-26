var express = require('express');
var app = express();

var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');
var session = require('express-session');
var moment = require('moment');
var querystring = require('querystring');

// Very basic HTML templates
var pages = require('./pages');
var authHelper = require('./authHelper');

// Configure express
// Set up rendering of static files
app.use(express.static('static'));
// Need JSON body parser for most API responses
app.use(bodyParser.json());
// Set up cookies and sessions to save tokens
app.use(cookieParser());
app.use(session(
  { secret: '0dc529ba-5051-4cd6-8b67-c9a901bb8bdf',
    resave: false,
    saveUninitialized: false 
  }));
  
// Home page
app.get('/', function(req, res) {
  res.send(pages.loginPage(authHelper.getAuthUrl()));
});

app.get('/authorize', function(req, res) {
  var authCode = req.query.code;
  if (authCode) {
    console.log('');
    console.log('Retrieved auth code in /authorize: ' + authCode);
    authHelper.getTokenFromCode(authCode, tokenReceived, req, res);
  }
  else {
    // redirect to home
    console.log('/authorize called without a code parameter, redirecting to login');
    res.redirect('/');
  }
});

function tokenReceived(req, res, error, token) {
  if (error) {
    console.log('ERROR getting token:'  + error);
    res.send('ERROR getting token: ' + error);
  }
  else {
    // save tokens in session
    req.session.access_token = token.token.access_token;
    req.session.refresh_token = token.token.refresh_token;
    req.session.email = authHelper.getEmailFromIdToken(token.token.id_token);
    res.redirect('/logincomplete');
  }
}

app.get('/logincomplete', function(req, res) {
  var access_token = req.session.access_token;
  var refresh_token = req.session.access_token;
  var email = req.session.email;
  
  if (access_token === undefined || refresh_token === undefined) {
    console.log('/logincomplete called while not logged in');
    res.redirect('/');
    return;
  }
  
  res.send(pages.loginCompletePage(email));
});

app.get('/refreshtokens', function(req, res) {
  var refresh_token = req.session.refresh_token;
  if (refresh_token === undefined) {
    console.log('no refresh token in session');
    res.redirect('/');
  }
  else {
    authHelper.getTokenFromRefreshToken(refresh_token, tokenReceived, req, res);
  }
});

app.get('/logout', function(req, res) {
  req.session.destroy();
  res.redirect('/');
});

// Start the server
var server = app.listen(3000, function() {
  var host = server.address().address;
  var port = server.address().port;
  
  console.log('Example app listening at http://%s:%s', host, port);
});