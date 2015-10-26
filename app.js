var express = require('express');
var app = express();

var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');
var session = require('express-session');
var moment = require('moment');

// Very basic HTML templates
var pages = require('./pages');

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
  res.send(pages.loginPage('#'));
});

// Start the server
var server = app.listen(3000, function() {
  var host = server.address().address;
  var port = server.address().port;
  
  console.log('Example app listening at http://%s:%s', host, port);
});