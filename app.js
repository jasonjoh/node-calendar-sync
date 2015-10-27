var express = require('express');
var app = express();

var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');
var session = require('express-session');
var moment = require('moment');
var querystring = require('querystring');
var outlook = require('node-outlook');

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

app.get('/sync', function(req, res) {
  var token = req.session.access_token;
  var email = req.session.email;
  if (token === undefined || email === undefined) {
    console.log('/sync called while not logged in');
    res.redirect('/');
    return;
  }
  
  // Set the endpoint to API v2
  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
  // Set the user's email as the anchor mailbox
  outlook.base.setAnchorMailbox(req.session.email);
  // Set the preferred time zone
  outlook.base.setPreferredTimeZone('Eastern Standard Time');
  
  // Use the syncUrl if available
  var requestUrl = req.session.syncUrl;
  if (requestUrl === undefined) {
    // Calendar sync works on the CalendarView endpoint
    requestUrl = outlook.base.apiEndpoint() + '/Me/CalendarView';
  }
  
  // Set up our sync window from midnight on the current day to
  // midnight 7 days from now.
  var startDate = moment().startOf('day');
  var endDate = moment(startDate).add(7, 'days');
  // The start and end date are passed as query parameters
  var params = {
    startDateTime: startDate.toISOString(),
    endDateTime: endDate.toISOString()
  };
  
  // Set the required headers for sync
  var headers = {
    Prefer: [ 
      // Enables sync functionality
      'odata.track-changes',
      // Requests only 5 changes per response
      'odata.maxpagesize=5'
    ]
  };
  
  var apiOptions = {
    url: requestUrl,
    token: token,
    headers: headers,
    query: params
  };
  
  outlook.base.makeApiCall(apiOptions, function(error, response) {
    if (error) {
      console.log(JSON.stringify(error));
      res.send(JSON.stringify(error));
    }
    else {
      if (response.statusCode !== 200) {
        console.log('API Call returned ' + response.statusCode);
        res.send('API Call returned ' + response.statusCode);
      }
      else {
        var nextLink = response.body['@odata.nextLink'];
        if (nextLink !== undefined) {
          req.session.syncUrl = nextLink;
        }
        var deltaLink = response.body['@odata.deltaLink'];
        if (deltaLink !== undefined) {
          req.session.syncUrl = deltaLink;
        }
        res.send(pages.syncPage(email, response.body.value));
      }
    }
  });
});

app.get('/viewitem', function(req, res) {
  var itemId = req.query.id;
  var access_token = req.session.access_token;
  var email = req.session.email;
  
  if (itemId === undefined || access_token === undefined) {
    res.redirect('/');
    return;
  }
  
  var select = {
    '$select': 'Subject,Attendees,Location,Start,End,IsReminderOn,ReminderMinutesBeforeStart'
  };
  
  var getEventParameters = {
    token: access_token,
    eventId: itemId,
    odataParams: select
  };
  
  outlook.calendar.getEvent(getEventParameters, function(error, event) {
    if (error) {
      console.log(error);
      res.send(error);
    }
    else {
      res.send(pages.itemDetailPage(email, event));
    }
  });
});

app.get('/updateitem', function(req, res) {
  var itemId = req.query.eventId;
  var access_token = req.session.access_token;
  
  if (itemId === undefined || access_token === undefined) {
    res.redirect('/');
    return;
  }
  
  var newSubject = req.query.subject;
  var newLocation = req.query.location;
  
  console.log('UPDATED SUBJECT: ', newSubject);
  console.log('UPDATED LOCATION: ', newLocation);
  
  var updatePayload = {
    Subject: newSubject,
    Location: {
      DisplayName: newLocation
    }
  };
  
  var updateEventParameters = {
    token: access_token,
    eventId: itemId,
    update: updatePayload
  };
  
  outlook.calendar.updateEvent(updateEventParameters, function(error, event) {
    if (error) {
      console.log(error);
      res.send(error);
    }
    else {
      res.redirect('/viewitem?' + querystring.stringify({ id: itemId }));
    }
  });
});

app.get('/deleteitem', function(req, res) {
  var itemId = req.query.id;
  var access_token = req.session.access_token;
  
  if (itemId === undefined || access_token === undefined) {
    res.redirect('/');
    return;
  }
  
  var deleteEventParameters = {
    token: access_token,
    eventId: itemId
  };
  
  outlook.calendar.deleteEvent(deleteEventParameters, function(error, event) {
    if (error) {
      console.log(error);
      res.send(error);
    }
    else {
      res.redirect('/sync');
    }
  });
});

// Start the server
var server = app.listen(3000, function() {
  var host = server.address().address;
  var port = server.address().port;
  
  console.log('Example app listening at http://%s:%s', host, port);
});