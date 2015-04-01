/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on.
 */
var ADDON_TITLE = 'CTracker';

/**
 * An admin email that is for a user who can configure add-on and enable 
 * it.
 */
var ADMIN_EMAIL = 'drendov@mymail.com';

/**
 * A global constant 'notice' text to include with each email
 * notification.
 */
var NOTICE = "CTracker was created as an add-on to track \
changes in a spreadsheet and notify an owner about changes. \
This is an experimental module. Collaborators using this add-on on \
the same spreadsheet will be able to look through revisions history, \
but will not be able to send report to anywhere else except \
emails that were set by admin of this add-on. \
                                                        OIE SUP team";

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Adds a custom menu to the active spreadsheet to show the add-on sidebars.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = SpreadsheetApp.getUi().createAddonMenu(); 
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    menu.addItem('Please, authorize add-on', 'showRevHistorySidebar');
    menu.addItem('About', 'showAbout');
  } else {
    menu.addItem('Revision history', 'showRevHistorySidebar');
    menu.addItem('Reports', 'showReportsSidebar');
    if (isAdmin()) {
      menu.addSeparator();
      menu.addItem('Configuration', 'showConfigurationSidebar');
    }
    menu.addSeparator();
    menu.addItem('About', 'showAbout');
   } 
  menu.addToUi();
}

/**
 * Opens a sidebar in the sheet containing the revisions history interface for
 * looking trough changes that have been done during the time when the add-on was active.
 */
function showRevHistorySidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('RevisionsSidebar')
      .setTitle('Revision history');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a sidebar in the sheet containing the reports interface for
 * making and sendig detailed reports about changes for some period of time.
 */
function showReportsSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('ReportsSidebar')
      .setTitle('Reports');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a sidebar in the sheet containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showConfigurationSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('ConfigurationSidebar')
      .setTitle('Add-on configuration');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setWidth(420)
      .setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(ui, 'About CTracker');
}

/**
 * Save sidebar settings to this form's Properties, and update the onEdit
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
  adjustCTrackerTrigger();
}

/**
 * Queries the User Properties and adds additional data required to populate
 * the sidebar UI elements.
 *
 * @return {Object} A collection of Property values and
 *     related data used to fill the configuration sidebar.
 */
function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getActiveUser().getEmail();
  }

  return settings;
}

/**
 * Is a user an admin fo the add-on.
 *
 * @return {Object} 'true' if an active user is an admin
 */
function isAdmin() {
  var isAdmin = false;
  
  if (ADMIN_EMAIL === Session.getActiveUser().getEmail()) {
    isAdmin = true;
  }
  
  return isAdmin;   
}

/**
 * Adjust the onEdit trigger based on user's requests.
 */
function adjustCTrackerTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  //FIXME: !Critical! Avoid multiple triggers for each user who is configuring the add-on
  var triggers = ScriptApp.getUserTriggers(sheet);
  var settings = PropertiesService.getDocumentProperties();
  var triggerNeeded =
      settings.getProperty('isEnabled') == 'true';

  // Create a new trigger if required; delete existing trigger
  //   if it is not needed.
  var existingTrigger = null;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_EDIT) {
      existingTrigger = triggers[i];
      break;
    }
  }
  if (triggerNeeded && !existingTrigger) {
    var trigger = ScriptApp.newTrigger('respondToEditSheet')
        .forSpreadsheet(sheet)
        .onEdit()
        .create();
    Logger.log('Trigger id = ' + trigger.getUniqueId() + ' set handler = ' + ScriptApp.getProjectTriggers()[0].getHandlerFunction());
    settings.setProperty('triggerId', trigger.getUniqueId());
  } else if (!triggerNeeded && existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}


/**
 * Responds to a form submission event if a respondToEditSheet trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
function respondToEditSheet(e) {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  // This check is required when using triggers with add-ons to maintain
  // functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to reauthorize; the normal trigger action is not
    // conducted, since it authorization needs to be provided first. Send at
    // most one 'Authorization Required' email a day, to avoid spamming users
    // of the add-on.
    sendReauthorizationRequest();
  } else {
    // All required authorizations has been granted, so continue to respond to
    // the trigger event.

    // Check if the owner needs to be notified and the add-on is enabled; if so, update revision histiry and send the notification.
    if (settings.getProperty('isEnabled') == 'true') {
       updateRevisionHistory();
    }

  }
}

/**
 * Make report about sheets and cells that have been changed
 * for some period of time
 *
 * @param {Object} from Timestamp from 
 * @param {Object} to Timestamp to
 *
 * @return {Object} HTML table with ready report
 */
function getReport(from, to) {
  var distinct = [];
  var body = '';
  var data = filterData(from, to);

  // FIXME: Hide table header if in history only one System Event record about erasing history logs
  if (data.length > 0) {
    body  = '<table class="report-table" id="report" cellspacing="0"><thead><tr>';
    body += '<th>Sheet</th>';
    body += '<th>Cell</th>';
    body += '</tr></thead><tbody>';
    for (var i = 0; i < data.length; i++) {
      body += '<tr>';
      body += '<td><a href="#" data-sheet="' + data[i].sheet + '" data-cell="' + data[i].cell + '" title="Changed in ' + data[i].date + '">' + data[i].sheet + '</a></td>';
      body += '<td>' + data[i].cell + '</td>';
      body += '</tr>';
    }
    body += '</tbody></table>';
  }

  return body;
}

/**
 * Filter distinct data from revision history
 * between the dates 'from' and 'to'
 *
 * @param {Object} from Timestamp from 
 * @param {Object} to Timestamp to
 *
 * @return {Object} Part of revision history data between the dates ordered desc
 */
function filterData(from, to) {
  var history = getStoreData("history");
  var distinct = [];
  var results = [];
  if (history.length > 0) {
    for (var i = history.length-1; i >= 0; i--) {
      var ts = new Date(history[i].timestamp).valueOf();
      if  (ts.toString() > from && ts.toString() < to) {
        var key = history[i].sheet + history[i].cell;
        if(distinct.indexOf(key) == -1 && history[i].sheet != 'System') {
          results.push(history[i]);
          distinct.push(key);
        }
      }
    }
  }
  return results;
}

/**
 * Send a report to users in creatorEmail property
 * with revision history detailed data for 
 * some period of time
 *
 * @param {Object} from Timestamp from 
 * @param {Object} to Timestamp to
 */
function sendUserReport(from, to) {
  var settings = PropertiesService.getDocumentProperties();
  var addresses = settings.getProperty('creatorEmail').split(',');
  var status = "";

  if (MailApp.getRemainingDailyQuota() > addresses.length && settings.getProperty('creatorNotify') == 'true') {
    var template = HtmlService.createTemplateFromFile('ReportNotification');
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    template.sheetName = ss.getName();
    template.sheetURL = ss.getUrl();  
    template.from = Utilities.formatDate(new Date(from), "GMT+04", "dd MMM, HH:mm");
    template.to = Utilities.formatDate(new Date(to), "GMT+04", "dd MMM, HH:mm");
    template.data = filterData(from, to);
    template.notice = NOTICE;
    
    var message = template.evaluate();
    MailApp.sendEmail(addresses,
          'Changes report',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    status = 'Successfully sent to ' + addresses + '. Remaining Daily Quota = ' + MailApp.getRemainingDailyQuota();

    //Browser.msgBox('Message: ' + status,  Browser.Buttons.OK);
  } else {
    status = 'Sorry, reports sending is disabled in the add-on configuration. Please, ping an admin of the plugin!';
  }
    
  return status;

}

/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template =
          HtmlService.createTemplateFromFile('AuthorizationEmail');
      template.url = authInfo.getAuthorizationUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(Session.getActiveUser().getEmail(),
          'Authorization Required',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
    settings.setProperty('lastAuthEmailDate', today);
  }
}

/**
 * Update revision history every time when a user edits 
 * cell and this change is saved.
 */
function updateRevisionHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties();
  
  var newcontent = ss.getActiveCell().getValue().toString();
  if (newcontent != "") {
    var cell = ss.getActiveCell().getA1Notation();
    var logentry = {"sheet": sheet.getName(),
                    "cell": ss.getActiveCell().getA1Notation(),
                    "timestamp": new Date().valueOf(),
                    "date": Utilities.formatDate(new Date(), "GMT+04", "dd MMM, HH:mm"),
                    "author": getOwnName(),
                    "email": Session.getEffectiveUser().getEmail(),
                    "content": newcontent
                   }
    
    var history = getStoreData("history");
    history.push(logentry)
    setStoreData("history", history); 
    
    // FIXME: Update revision history if this sidebar is open
    //var app = UiApp.getActiveApplication();
    //google.script.host.reloadLogs();
    
  }
  
}

/**
 * Get stored revision history from DocumentProperties
 *
 * @return {Object} A collection of Property values 
 *                  that contains revision history
 */
function getLogs(){
  var logs = getStoreData("history");
  return logs;
}

/**
 * Reset saved revision history from DocumentProperties
 * and saved who did that
 * To-Do: Allow clear history only for EPAM team only
 */
function resetStoreData(){
  var resetarray = [];
  if (isAdmin()) {
    resetarray.push({
      "sheet": "System",
      "cell": "event",
      "timestamp": new Date().valueOf(),
      "date": Utilities.formatDate(new Date(), "GMT+04", "dd MMM, HH:mm"),
      "author": getOwnName(),
      "email": Session.getEffectiveUser().getEmail(),
      "content": '<span style="color:red;font-weight:bold;">Revision history has been cleared!</span>'
    });
    setStoreData("history", resetarray);
  } else {
    Browser.msgBox('Sorry, only admins can clear Revision history!',  Browser.Buttons.OK);
  }  
  return resetarray;
}


/**
 * Store data at some point in the application    
 *
 * @param {Object} storageName A name of propery
 * @param {Object} data2store JSON object to store
 */
function setStoreData(storageName, data2store){   
  PropertiesService   
   .getDocumentProperties()   
   .setProperty(storageName, JSON.stringify(data2store));  
}   
    
/**
 * Read stored data as JSON at some point in the application    
 *
 * @param {Object} storageName A name of propery
 * @return {Object} A collection of Property values
 */
function getStoreData(storageName){   
  var data = JSON.parse(PropertiesService   
                  .getDocumentProperties()   
                  .getProperty(storageName)   
                 );   
  return data;    
}   

/**
 * Get current user's name, by accessing their contacts.
 *
 * @returns {String} First name (GivenName) if available,
 *                   else FullName, or login ID (userName)
 *                   if record not found in contacts.
 */
function getOwnName(){
  var email = Session.getActiveUser().getEmail();
  var self = ContactsApp.getContact(email);

  // If user has themselves in their contacts, return their name
  if (self) {
    // Prefer given name, if that's available
    var name = self.getFullName();
    // But we will settle for the full name
    if (!name) name = self.getGivenName();
    return name;
  }
  // If they don't have themselves in Contacts, return the bald userName.
  else {
    var userName = Session.getEffectiveUser().getUsername();
    return userName;
  }
}

function getEmailBodyForError (error) {
  
  var body = "message : " + error.message + " " + error.stack;
  body = body + "file :" + error.fileName + "\n";
  body = body + "line :" + error.lineNumber + "\n";
  body = body + "\n";
  
  return body;
}

/**
 * FIXME: Doesn't work
 */
function doGet(request) {
  return HtmlService.createTemplateFromFile('Page')
      .evaluate();
}

/**
 * FIXME: Doesn't work
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
