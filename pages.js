var querystring = require('querystring');

var baseHtml = '<html>' +
  '<head>' +
    '<meta content="IE=edge" http-equiv="X-UA-Compatible">' +
    '<meta charset="utf-8">' +
    '<title>%title%</title>'  +
    '<link type="text/css" rel="stylesheet" href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">' +
    '<link type="text/css" rel="stylesheet" href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">' +
    '<link type="text/css" rel="stylesheet" href="styles/app.css">' +
  '</head>' +
  '<body>' +
    '<div id="main-content" class="ms-Grid">' +
      '<div class="ms-Grid-row">' + 
        '<div id="title-banner" class="ms-font-su">Outlook Calendar Sync Demo</div>' +
      '</div>' +
      '<div id="body-content" class="ms-Grid-row">' +
        '%body%' +
      '</div>' + 
    '</div>' + 
  '</body>' +
'</html>';

var buttonRow = '<div class="ms-Grid-row">' +
    '<div id="user-email" class="ms-font-l">Signed in as: %email%</div>' +
  '</div>' +
  '<div class="ms-Grid-row">' +
    '<div class="ms-Grid-col ms-u-sm4">' +
      '<a class="ms-Button ms-Button--primary" href="/sync"><span class="ms-Button-label">Sync calendar</span></a>' +
    '</div>' +
    '<div class="ms-Grid-col ms-u-sm4">' +
      '<a class="ms-Button ms-Button--primary" href="/refreshtokens"><span class="ms-Button-label">Refresh tokens</span></a>' +
    '</div>' +
    '<div class="ms-Grid-col ms-u-sm4">' +
      '<a class="ms-Button ms-Button--primary" href="/logout"><span class="ms-Button-label">Logout</span></a>' +
    '</div>' +
  '</div>';
  
function extractId(change) {
  return change.id.match(/'([^']+)'/)[1];
}

function getViewItemLink(change) {
  if (change.reason && change.reason === 'deleted') {
    return '';
  }
  
  var link = '<a href="/viewitem?';
  link += querystring.stringify({ id: change.Id });
  link += '">View Item</a>';
  return link;
}

function getAttendeesStrings(attendees) {
  var displayStrings = {
    required: '',
    optional: '',
    resources: ''
  };
  
  attendees.forEach(function(attendee) {
    var attendeeName = (attendee.EmailAddress.Name === undefined) ? 
      attendee.EmailAddress.Address : attendee.EmailAddress.Name;
    switch (attendee.Type) {
      // Required
      case "Required":
        if (displayStrings.required.length > 0) {
          displayStrings.required += '; ' + attendeeName;
        }
        else {
          displayStrings.required += attendeeName;
        }
        break;
      // Optional
      case "Optional":
        if (displayStrings.optional.length > 0) {
          displayStrings.optional += '; ' + attendeeName;
        }
        else {
          displayStrings.optional += attendeeName;
        }
        break;
      // Resources
      case "Resource":
        if (displayStrings.resources.length > 0) {
          displayStrings.resources += '; ' + attendeeName;
        }
        else {
          displayStrings.resources += attendeeName;
        }
        break;
    }
  });
  
  return displayStrings;
}

module.exports = {
  loginPage: function(signinUrl) {
    var html = '<a id="signin-button" class="ms-Button ms-Button--primary" href="' + signinUrl + '"><span class="ms-Button-label">Click here to sign in</span></a>';
    
    return baseHtml.replace('%title%', 'Login').replace('%body%', html);
  },
  
  loginCompletePage: function(userEmail) {
    var html = '<div class="ms-Grid">';
    html += buttonRow.replace('%email%', userEmail);
    html += '</div>';
    
    return baseHtml.replace('%title%', 'Main').replace('%body%', html);
  },
  
  syncPage: function(userEmail, changes) {
    var html = '<div class="ms-Grid">';
    html += buttonRow.replace('%email%', userEmail);
    
    html += '<div id="table-row" class="ms-Grid-row">';
    html += '  <div class="ms-font-l ms-fontWeight-semibold">Changes</div>';
    html += '  <div class="ms-Table">';
    html += '    <div class="ms-Table-row">';
    html += '      <div class="ms-Table-cell">Change type</div>';
    html += '      <div class="ms-Table-cell">Details</div>';
    html += '      <div class="ms-Table-cell"></div>';
    html += '    </div>';
    
    if (changes && changes.length > 0) {
      changes.forEach(function(change){
        var changeType = (change.reason && change.reason === 'deleted') ? 'Delete' : 'Add/Update';
        var detail = (changeType === 'Delete') ? extractId(change) : change.Subject;
        html += '<div class="ms-Table-row">';
        html += '  <div class="ms-Table-cell">' + changeType + '</div>';
        html += '  <div class="ms-Table-cell">' + detail + '</div>';
        html += '  <div class="ms-Table-cell">' + getViewItemLink(change) + '</div>';
        html += '</div>';
      });
    }
    else {
      html += '<div class="ms-Table-row"><div class="ms-Table-cell">-</div><div class="ms-Table-cell">No Changes</div></div>';
    }
    
    html += '  </div>';
    html += '</div>';
    
    html += '<pre>' + JSON.stringify(changes, null, 2) + '</pre>';
    return baseHtml.replace('%title%', 'Sync').replace('%body%', html);
  },
  
  itemDetailPage: function(userEmail, event) {
    var html = '<div class="ms-Grid">';
    html += buttonRow.replace('%email%', userEmail);
    
    html += '<form action="/updateitem" method="get">';
    
    html += '<input name="eventId" type="hidden" value="' + event.Id + '"/>';
    
    html += '<div id="event-subject" class="ms-Grid-row">';
    html += '  <div class="ms-Grid-col ms-u-sm12">';
    html += '    <div class="ms-TextField">';
    html += '      <label class="ms-Label">Subject</label>';
    html += '      <input name="subject" class="ms-TextField-field" value="' + event.Subject + '"/>';
    html += '    </div>';
    html += '  </div>';
    html += '</div>';
    
    html += '<div class="ms-Grid-row">';
    html += '  <div class="ms-Grid-col ms-u-sm12">';
    html += '    <div class="ms-TextField">';
    html += '      <label class="ms-Label">Location</label>';
    html += '      <input name="location" class="ms-TextField-field" value="' + event.Location.DisplayName + '"/>';
    html += '    </div>';
    html += '  </div>';
    html += '</div>';
    
    if (event.IsReminderOn) {
      html += '<div class="ms-Grid-row">';
      html += '  <div class="ms-Grid-col ms-u-sm12">';
      html += '    <div class="ms-TextField is-disabled">';
      html += '      <label class="ms-Label">Reminder minutes before start</label>';
      html += '      <input class="ms-TextField-field" value="' + event.ReminderMinutesBeforeStart + '"/>';
      html += '    </div>';
      html += '  </div>';
      html += '</div>';
    }
    
    var attendees = getAttendeesStrings(event.Attendees);
    
    if (attendees.required.length > 0) {
      html += '<div class="ms-Grid-row">';
      html += '  <div class="ms-Grid-col ms-u-sm12">';
      html += '    <div class="ms-TextField is-disabled">';
      html += '      <label class="ms-Label">Required attendees</label>';
      html += '      <input class="ms-TextField-field" value="' + attendees.required + '"/>';
      html += '    </div>';
      html += '  </div>';
      html += '</div>';
    }
    
    if (attendees.optional.length > 0) {
      html += '<div class="ms-Grid-row">';
      html += '  <div class="ms-Grid-col ms-u-sm12">';
      html += '    <div class="ms-TextField is-disabled">';
      html += '      <label class="ms-Label">Optional attendees</label>';
      html += '      <input class="ms-TextField-field" value="' + attendees.optional + '"/>';
      html += '    </div>';
      html += '  </div>';
      html += '</div>';
    }
    
    if (attendees.resources.length > 0) {
      html += '<div class="ms-Grid-row">';
      html += '  <div class="ms-Grid-col ms-u-sm12">';
      html += '    <div class="ms-TextField is-disabled">';
      html += '      <label class="ms-Label">Resources</label>';
      html += '      <input class="ms-TextField-field" value="' + attendees.resources + '"/>';
      html += '    </div>';
      html += '  </div>';
      html += '</div>';
    }
    
    html += '<div class="ms-Grid-row">';
    html += '  <div class="ms-Grid-col ms-u-sm6">';
    html += '    <div class="ms-TextField is-disabled">';
    html += '      <label class="ms-Label">Start</label>';
    html += '      <input class="ms-TextField-field" value="' + new Date(event.Start.DateTime).toString() + '"/>';
    html += '    </div>';
    html += '  </div>';
    html += '  <div class="ms-Grid-col ms-u-sm6">';
    html += '    <div class="ms-TextField is-disabled">';
    html += '      <label class="ms-Label">End</label>';
    html += '      <input class="ms-TextField-field" value="' + new Date(event.End.DateTime).toString() + '"/>';
    html += '    </div>';
    html += '  </div>';
    html += '</div>';
    
    html += '<div id="action-buttons" class="ms-Grid-row">';
    html += '  <div class="ms-Grid-col ms-u-sm6">';
    html += '    <input type="submit" class="ms-Button ms-Button--primary ms-Button-label" value="Update item"/>';
    html += '  </div>';
    html += '  <div class="ms-Grid-col ms-u-sm6">';
    html += '    <a class="ms-Button ms-Button--primary" href="/deleteitem?' + querystring.stringify({ id: event.Id }) + '"><span class="ms-Button-label">Delete item</span></a>';
    html += '  </div>';
    html += '</div>';
    html += '</form>';
    
    html += '<pre>' + JSON.stringify(event, null, 2) + '</pre>';
    // end grid
    html += '</div>';
    
    return baseHtml.replace('%title%', event.Subject).replace('%body%', html);
  }
};