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
 * A global constant 'notice' text to include with each email
 * notification.
 */
var NOTICE = "CTracker was created as an add-on to check \
changes in an sheet and notify an owner about changes. \
This is an experimental module. Collaborators using this add-on on \
the same form will be able to adjust the notification settings, but will not be \
able to disable the notification triggers set by other collaborators.";


/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
      .addItem('Configure notifications', 'showSidebar')
      .addItem('About', 'showAbout')
      .addSeparator()
      .addItem('Revision history', 'showRevHistorySidebar')
      .addItem('Reports', 'showReportsSidebar')
      .addToUi();
}

function showRevHistorySidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('RevisionsSidebar')
      .setTitle('CTracker Revisions history');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showReportsSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('ReporstSidebar')
      .setTitle('CTracker Reports');
  SpreadsheetApp.getUi().showSidebar(ui);
}

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
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Change Tracker');
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
 * Save sidebar settings to this form's Properties, and update the onFormSubmit
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings) {

  // If history container hasn't been initializated yet.
  if (!settings.history) {
    settings.history = [];
  }
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
  var history = settings.history;
  var history1 = settings.history[0];
  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getEffectiveUser().getEmail();
  }

  return settings;
}

/**
 * Adjust the onFormSubmit trigger based on user's requests.
 */
function adjustCTrackerTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
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
    Logger.log(ScriptApp.getProjectTriggers()[0].getHandlerFunction());
  } else if (!triggerNeeded && existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}


/**
 * Responds to a form submission event if a onFormSubmit trigger has been
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
    if (settings.getProperty('isEnabled') == 'true' && settings.getProperty('creatorNotify') == 'true' && MailApp.getRemainingDailyQuota() > 0) {
       updateRevisionHistory();
    }

  }
}

/**
 *
 * 
 */
function getReport(from, to) {
  var html2output = "";
  var history = getStoreData("history");
  var distinct = [];
  //"Tracker settingsD4", "Tracker settingsD7", "Tracker settingsG5", "Tracker settingsH5", "Tracker settingsH6"];
  //var key = "Tracker settingsH6";
  var from = 1426366800000;
  var to = 1426536000000;
  //Browser.msgBox('from: ' + from + ' to: ' + to,  Browser.Buttons.OK);  
  
  for (var i = history.length-1; i > 0; i--) {
    var ts = new Date(history[i].timestamp).valueOf().toString();
    if  (ts > from && ts < to) {
      var key = history[i].sheet + history[i].cell;
      if(distinct.indexOf(key) == -1) {
        html2output += '<li>[' + history[i].sheet + ']:' + history[i].cell + ' = ' + history[i].content + '</li>';
        distinct.push(key);
      }
    }
  }

  //Browser.msgBox('html2output: ' + html2output,  Browser.Buttons.OK);
  return html2output;
}

/**
 *
 * 
 */
function sendUserReport() {
  var status = "";
  
  status = "OK";
  
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
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
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
 * Sends out creator notification email(s) if the current number
 * of form responses is an even multiple of the response step
 * setting.
 *
 * FIXME: Need to rework this module
 *
 */
function updateRevisionHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties();
  var responseStep = settings.getProperty('responseStep');
  responseStep = responseStep ? parseInt(responseStep) : 10;

  // If the total number of form responses is an even multiple of the
  // response step setting, send a notification email(s) to the form
  // creator(s). For example, if the response step is 10, notifications
  // will be sent when there are 10, 20, 30, etc. total form responses
  // received.
  
  var addresses = settings.getProperty('creatorEmail').split(',');
  var shhetname = sheet.getName();
  var newcontent = ss.getActiveCell().getValue().toString();
  if (newcontent != "") {
    var cell = ss.getActiveCell().getA1Notation();
    var logentry = {"sheet": sheet.getName(),
                    "cell": ss.getActiveCell().getA1Notation(),
                    "timestamp": new Date().valueOf(),
                    "date": Utilities.formatDate(new Date(), "GMT", "dd MMM, HH:mm"),
                    "author": getOwnName(),
                    "email": Session.getEffectiveUser().getEmail(),
                    "content": newcontent
                   }
    
    var history = getStoreData("history");
    history.push(logentry)
    setStoreData("history", history); 
    
    var app = UiApp.getActiveApplication();
    google.script.host.reloadLogs();

    //Browser.msgBox('History: ' + settings.getProperty("history"),  Browser.Buttons.OK);
    //Browser.msgBox('logentry: ' + logentry,  Browser.Buttons.OK);
    
  }

  /*if (form.getResponses().length % responseStep == 0) {
    if (MailApp.getRemainingDailyQuota() > addresses.length) {
      var template =
          HtmlService.createTemplateFromFile('CreatorNotification');
      template.summary = form.getSummaryUrl();
      template.responses = form.getResponses().length;
      template.title = form.getTitle();
      template.responseStep = responseStep;
      template.formUrl = form.getEditUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(settings.getProperty('creatorEmail'),
          form.getTitle() + ': Form submissions detected',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
      Logger.log("Remaining Daily Quota = " + MailApp.getRemainingDailyQuota());
    }
  }*/
  
}

/**
 * Get stored revision history from DocumentProperties
 *
 * @return {Object} A collection of Property values 
 *                  that contains revision history
 */
function getLogs(){
  var logs = getStoreData("history")
  return logs;
}

/**
 * Reset saved revision histiry from DocumentProperties
 *
 * TO-DO: Added info who reset logs
 */
function resetStoreData(){
  var emptyarr = [];
  setStoreData("history", emptyarr);
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
  var email = Session.getEffectiveUser().getEmail();
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
