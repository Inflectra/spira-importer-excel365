/*
 *
 * =============
 * GENERAL SETUP
 * =============
 *
 */

//Global variable to decide which application is being used
var IS_GOOGLE = typeof UrlFetchApp != "undefined";

/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 * Please comment/uncomment this block of code for Google Sheets/Excel*/


export {
  clearAll,
  error,
  getBespoke,
  getComponents,
  getCustoms,
  getFromSpiraExcel,
  getProjects,
  getProjectTemplates,
  getTemplateLists,
  getReleases,
  getUsers,
  getTemplateFromProjectId,
  operationComplete,
  sendToSpira,
  templateLoader,
  warn,
  isUserAdmin
};

import { showPanel, hidePanel } from './taskpane.js';

import { params } from './model.js';

// ===END OF EXCEL SPECIFIC CODE===

/*
 *
 * ============================
 * GOOGLE SHEETS SPECIFIC SETUP
 * ============================
 * Please uncomment/comment this block of code for Google Sheets/Excel
 */

//Import Model (data.js.html)
/*var importModel = HtmlService.createHtmlOutputFromFile('data.js').getContent();
eval(importModel.replace('<script>', '').replace('</script>', ''));*/

/* ===END OF GOOGLE SPECIFIC CODE===


/*
 * =======
 * TODO
 * =======
 
 - make sure when you change project / art the get / send buttons are disabled
 - check what happens when add more rows from get than are on sheet. Does the validation get copied down?
 */

// globals
var API_BASE = '/services/v6_0/RestService.svc/',
  API_PROJECT_BASE = '/services/v6_0/RestService.svc/projects/',
  API_PROJECT_BASE_NO_SLASH = '/services/v6_0/RestService.svc/projects',
  API_TEMPLATE_BASE = '/services/v6_0/RestService.svc/project-templates/',
  API_TEMPLATE_BASE_NO_SLASH = '/services/v6_0/RestService.svc/project-templates',
  API_USER_BASE = '/services/v6_0/RestService.svc/users/usernames/',
  ART_ENUMS = {
    requirements: 1,
    testCases: 2,
    incidents: 3,
    releases: 4,
    testRuns: 5,
    tasks: 6,
    testSteps: 7,
    testSets: 8,
    risks: 14,
    folders: 114,
    components: 99,
    users: 98,
    customLists: 97,
    customValues: 96,
  },
  INITIAL_HIERARCHY_OUTDENT = -20,
  GET_PAGINATION_SIZE = 100,
  EXCEL_MAX_ROWS_PER_PAGE = 2000,
  ARTIFACT_MAX_PAGES = 500,
  FIELD_MANAGEMENT_ENUMS = {
    all: 1,
    standard: 2,
    subType: 3
  },
  STATUS_ENUM = {
    allSuccess: 1,
    someError: 2,
    allError: 3,
    wrongSheet: 4,
    existingEntries: 5,
    advancedMode:6,
    noData:7
  },
  SUBTYPE_IDS = ["TestCaseId", "TestStepId"],
  STATUS_MESSAGE_GOOGLE = {
    1: "All done! To send more data over, clear the sheet first.",
    2: "Sorry, but there were some problems (see the cells marked in red). Check any notes on the relevant ID field for explanations.",
    3: "We're really sorry, there was a problem sending data to Spira. Some records may not have been sent correctly. Please check in Spira to confirm.",
    4: "You are not on the correct worksheet. Please go to the sheet that matches the one listed on the Spira taskpane / the selection you made in the sidebar.",
    5: "Some/all of the rows already exist in SpiraPlan. These rows have not been re-added.",
    6: "Data sent! Since you are using the advanced fields mode, any errors related to comments and/or associations will be marked as red in the spreadsheet. To send more data over, clear the sheet first.",
    7: "No records were returned from Spira. Please verify the page number, product, and artifact. If the problem persists, contact your Spira administrator."
  },
  STATUS_MESSAGE_EXCEL = {
    1: "All done! To send more data over, clear the sheet first.",
    2: "Sorry, but there were some problems (see the cells marked in red). Check any notes on the relevant ID field for explanations.",
    3: "We're really sorry, there was a problem sending data to Spira. Some records may not have been sent correctly. Please check in Spira to confirm.",
    4: "You are not on the correct worksheet. Please go to the sheet that matches the one listed on the Spira taskpane / the selection you made in the sidebar.",
    5: "Some/all of the rows already exist in SpiraPlan. These rows have not been re-added.",
    6: "Data sent! Since you are using the advanced fields mode, any errors related to comments and/or associations will be marked as red in the spreadsheet. To send more data over, clear the sheet first.",
    7: "No records were returned from Spira. Please verify the page number, product, and artifact. If the problem persists, contact your Spira administrator."
  },
  CUSTOM_PROP_TYPE_ENUM = {
    //internal-handling only
    101: "StringValue",
    102: "IntegerValue",
    103: "DecimalValue",
    104: "BooleanValue",
    105: "DateTimeValue",
    106: "IntegerValue",
    107: "IntegerListValue",
    108: "IntegerValue",
    //Spira ID values from API
    1: "StringValue",
    2: "IntegerValue",
    3: "DecimalValue",
    4: "BooleanValue",
    5: "DateTimeValue",
    6: "IntegerValue",
    7: "IntegerListValue",
    8: "IntegerValue",
    9: "StringValue",
    10: "IntegerValue",
    11: "DateTimeValue",
    12: "IntegerValue"
  },
  INLINE_STYLING = "style='font-family: sans-serif'",
  ART_PARENT_IDS = {
    2: 'TestCaseId',
    7: 'TestCaseId',
    97: 'CustomPropertyListId',
    96: 'CustomPropertyListId'
  };

const EXCEL_NUMBER_OF_ROWS = 1048576;
const DAYS_BETWEEN_1900_1970 = 25567 + 2;

/*
 * ======================
 * INITIAL LOAD FUNCTIONS
 * ======================
 *
 * These functions are needed for initialization
 * All Google App Script (GAS) files are bundled by the engine
 * at start up so any non-scoped variables declared will be available globally.
 *
 */

// Google App script boilerplate install function
// opens app on install
function onInstall(e) {
  onOpen(e);
}



// App script boilerplate open function
// opens sidebar
// Method `addItem`  is related to the 'Add-on' menu items. Currently just one is listed 'Start' in the dropdown menu
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Start', 'showSidebar').addToUi();
}



// side bar function gets index.html and opens in side window
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('SpiraPlan by Inflectra');

  SpreadsheetApp.getUi().showSidebar(ui);
}



// This function is part of the google template engine and allows for modularization of code
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}









/*
 *
 * ========================
 * TEMPLATE PANEL FUNCTIONS
 * ========================
 *
 */

// copy the first sheet into a new sheet in the same spreadsheet
function save() {
  // pop up telling the user that their data will be saved
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will save the current sheet in a new sheet on this spreadsheet. Continue?', ui.ButtonSet.YES_NO);

  // returns with user choice
  if (response == ui.Button.YES) {
    // get first sheet of  active spreadsheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getActiveSheet();

    // get entire open spreadsheet id
    var id = spreadSheet.getId();

    // set current spreadsheet file as destination
    var destination = SpreadsheetApp.openById(id);

    // copy sheet to current spreadsheet in new sheet
    sheet.copyTo(destination);

    // returns true to queue success popup
    return true;
  } else {
    // returns false to ignore success popup
    return false;
  }
}

//Function to check if a given sheet name exists or not in the workbook 
async function checkSheetExists(sheetName) {
  Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
      .then(function () {
        sheets.items.forEach(function (sheet) {
          if (sheet.name == sheetName) {
            return true;
          }
        });
      });
  });
  return false;
}

//clears active sheet in spreadsheet
function clearAll(model) {

  if (IS_GOOGLE) {
    // get active spreadsheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = spreadSheet.getActiveSheet(),
      lastColumn = sheet.getLastColumn(),
      lastRow = sheet.getLastRow();

    //minimum requirements
    if (lastColumn == 0) { lastColumn = 1 };
    if (lastRow == 0) { lastRow = 1 };

    // Reset sheet name
    sheet.setName(new Date().getTime());
    sheet.clear();

    // clears data validations and notes from the entire sheet
    var range = sheet.getRange(1, 1, sheet.getMaxRows() - 1, lastColumn);
    range.clearDataValidations();
    range.clearDataValidations().clearNote();

    //unhide columns
    sheet.showColumns(1, lastColumn);

    // remove any protections on the sheet
    var protections = spreadSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.canEdit()) {
        protection.remove();
      }
    }
    //check if the database sheet exists 
    var dataBaseSheetName = createDatabaseSheetName(params.dataSheetName, model.currentProject.id, model.currentArtifact.id);

    var databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataBaseSheetName);

    if (databaseSheet != null) {
      //if we have a database worksheet, clear it 
      databaseSheet.clear();
    }
    else {
      //if we don't have a database worksheet, create one 
      var newDataBaseSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(dataBaseSheetName);
      newDataBaseSheet.hideSheet();
    }

    return true;
  } else {
    return Excel.run(context => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      // for excel we do not reset the sheet name because this can cause timing problems on some versions of Excel 
      sheet.getRange().clear();

      //check if the database sheet exists 
      var sheets = context.workbook.worksheets;
      sheets.load("items/name");

      var isDatabaseSheet = false;
      var dataBaseSheetName = createDatabaseSheetName(params.dataSheetName, model.currentProject.id, model.currentArtifact.id);

      return context.sync()
        .then(function () {
          sheets.items.forEach(function (singleSheet) {
            if (singleSheet.name == dataBaseSheetName) {
              isDatabaseSheet = true;
            }
          });

          if (!isDatabaseSheet) {
            //if we don't have a database worksheet, create one 
            var dbSheet = sheets.add(dataBaseSheetName);
            dbSheet.visibility = Excel.SheetVisibility.hidden;
          }
          else {
            //if we have a database worksheet, clear it 
            var worksheet = context.workbook.worksheets.getItemOrNullObject(dataBaseSheetName);
            worksheet.getRange().clear();
          }
          return context.sync();
        });
    })
  }
}


// handles showing popup messages to user
// @param: message - strng of the raw message to show user
// @param: messageTitle - strng of the message title to use
// @param: isTemplateLoadFail - bool about whether this message means that the template load sequence has failed
function popupShow(message, messageTitle, isTemplateLoadFail) {
  if (!message) return;
  else {
    if (IS_GOOGLE) {
      // Display a modal dialog box with custom HtmlService content.
      var htmlMessage = HtmlService.createHtmlOutput('<p ' + INLINE_STYLING + '>' + message + '</p>').setWidth(200).setHeight(75);
      SpreadsheetApp.getUi().showModalDialog(htmlMessage, messageTitle || "");
    } else {
      showPanel("confirm");
      document.getElementById("message-confirm").innerHTML = (messageTitle ? "<b>" + messageTitle + ":</b> " : "") + message;
      document.getElementById("btn-confirm-cancel").style.visibility = "hidden";
      document.getElementById("btn-confirm-ok").onclick = function () { popupHide() };
      return !isTemplateLoadFail ? null : {
        isTemplateLoadFail: isTemplateLoadFail,
        message: message
      };
    }
  }
}

function popupHide() {
  hidePanel("confirm");
  document.getElementById("message-confirm").innerHTML = "";
  document.getElementById("btn-confirm-cancel").style.visibility = "visible";
}



/*
 *
 * ====================
 * DATA "GET" FUNCTIONS
 * ====================
 *
 * functions used to retrieve data from Spira - things like projects and users, not specific records
 *
 */

// General fetch function, using Google's built in fetch api
// @param: currentUser - user object storing login data from client
// @param: fetcherUrl - url string passed in to connect with Spira
function fetcher(currentUser, fetcherURL) {
  //use google's Utilities to base64 decode if present, otherwise use standard JS (ie for MS Excel)
  var decoded = typeof Utilities != "undefined" ? Utilities.base64Decode(currentUser.api_key) : atob(currentUser.api_key);
  var APIKEY = typeof Utilities != "undefined" ? Utilities.newBlob(decoded).getDataAsString() : decoded;

  //build URL from args
  var fullUrl = currentUser.url + fetcherURL + "username=" + currentUser.userName + APIKEY;
  //set MIME type
  var params = { "Content-Type": "application/json", "accepts": "application/json" };

  //call Google fetch function (UrlFetchApp) if using google
  if (IS_GOOGLE) {
    try {
      var response = UrlFetchApp.fetch(fullUrl, params);
      //returns parsed JSON
      //unparsed response contains error codes if needed
      return JSON.parse(response);
    }
    catch (err) {
      error('network', err);
      throw 'network error!';
    }
    //for v6 API in Spira you HAVE to send a Content-Type header
  } else {
    return superagent
      .get(fullUrl)
      .set("Content-Type", "application/json", "accepts", "application/json")
  }

}


// Gets projects accessible by current logged in user
// This function is called on initial log in
// @param: currentUser - object with details about the current user
async function getProjects(currentUser) {
  var fetcherURL = API_PROJECT_BASE_NO_SLASH + '?';
  return fetcher(currentUser, fetcherURL);
}


// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
// @param: currentUser - object with details about the current user
async function getProjectTemplates(currentUser) {
  var fetcherURL = API_TEMPLATE_BASE_NO_SLASH + '?';
  return fetcher(currentUser, fetcherURL);
}

// Gets lists of selected template
// This function is called as soon as a template is selected on its dropdown
// @param: currentTemplate - object with details about the current Template
async function getTemplateLists(currentTemplate, currentUser) {
  var fetcherURL = API_TEMPLATE_BASE + currentTemplate + '/custom-lists?';
  return fetcher(currentUser, fetcherURL);
}


// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getTemplateFromProjectId(currentUser, projectId) {
  var fetcherURL = API_PROJECT_BASE + projectId + '?';
  return fetcher(currentUser, fetcherURL);
}

// Checks if the current user is a system administrator
// This function is called on initial log in
// @param: currentUser - object with details about the current user
async function isUserAdmin(currentUser) {
  var fetcherURL = API_USER_BASE + currentUser.userName + '?include_inactive=false&';
  return fetcher(currentUser, fetcherURL);
}


// Gets components for selected project.
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getComponents(currentUser, projectId) {
  var fetcherURL = API_PROJECT_BASE + projectId + '/components?active_only=true&include_deleted=false&';
  return fetcher(currentUser, fetcherURL);
}



// Gets custom fields for selected project and artifact
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: artifactName - int of the current artifact - API refers to this as the artifactTypeName but the ID is required
function getCustoms(currentUser, templateId, artifactName) {
  var fetcherURL = API_TEMPLATE_BASE + templateId + '/custom-properties/' + artifactName + '?';
  return fetcher(currentUser, fetcherURL);
}

// Gets data for a bespoke specified field (for selected project and artifact)
// @param: currentUser - object with details about the current user
// @param: templateId - int id for current template
// @param: projectId - int id for current project
// @param: artifactName - string name of the current artifact
// @param: field - object of the field from the model
function getBespoke(currentUser, templateId, projectId, artifactName, field) {
  var fetcherURL = "";
  // a couple of dynamic fields are project based - like folders
  if (field.bespoke.isProjectBased) {
    fetcherURL = API_PROJECT_BASE + projectId + field.bespoke.url + '?';
  } else if (field.bespoke.isSystemWide) {
    fetcherURL = API_BASE + field.bespoke.url + '?';
  }
  else {
    fetcherURL = API_TEMPLATE_BASE + templateId + field.bespoke.url + '?';
  }
  var results = fetcher(currentUser, fetcherURL);

  if (IS_GOOGLE) {
    return {
      artifactName: artifactName,
      field: field,
      values: results
    }
  } else {
    return results;
  }
}



// Gets releases for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getReleases(currentUser, projectId) {
  var fetcherURL = API_PROJECT_BASE + projectId + '/releases?';
  return fetcher(currentUser, fetcherURL);
}



// Gets users for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getUsers(currentUser, projectId) {
  var fetcherURL = API_PROJECT_BASE + projectId + '/users?';
  return fetcher(currentUser, fetcherURL);
}


function getArtifacts(user, projectId, artifactTypeId, startRow, numberOfRows, artifactId, templateId) {
  var fullURL = API_PROJECT_BASE + projectId;
  var response = null;
  switch (artifactTypeId) {
    case ART_ENUMS.requirements:
      fullURL += "/requirements?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=RequirementId&sort_direction=ASC&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.testCases:
      fullURL += "/test-cases?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=TestCaseId&sort_direction=ASC&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.testSteps:
      if (artifactId) {
        fullURL += "/test-cases/" + artifactId + "/test-steps?&";
        response = fetcher(user, fullURL);
      }
      break;
    case ART_ENUMS.incidents:
      fullURL += "/incidents/search?start_row=" + startRow + "&number_rows=" + numberOfRows + "&sort_field=Name&sort_direction=ASC&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.releases:
      fullURL += "/releases/search?start_row=" + startRow + "&number_rows=" + numberOfRows + "&sort_field=ReleaseId&sort_direction=ASC&";
      var rawResponse = poster("", user, fullURL);
      response = IS_GOOGLE ? JSON.parse(rawResponse) : rawResponse; // this particular return needs to be parsed here
      break;
    case ART_ENUMS.tasks:
      fullURL += "/tasks?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=TaskId&sort_direction=ASC&";
      var rawResponse = poster("", user, fullURL);
      response = IS_GOOGLE ? JSON.parse(rawResponse) : rawResponse; // this particular return needs to be parsed here
      break;
    case ART_ENUMS.risks:
      fullURL += "/risks?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=RiskId&sort_direction=ASC&";
      var rawResponse = poster("", user, fullURL);
      response = IS_GOOGLE ? JSON.parse(rawResponse) : rawResponse; // this particular return needs to be parsed here
      break;
    case ART_ENUMS.testSets:
      fullURL += "/test-sets?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=TestSetId&sort_direction=ASC&";
      if (IS_GOOGLE) {
        var rawResponse = JSON.stringify(fetcher(user, fullURL));
        response = JSON.parse(rawResponse); // this particular return needs to be parsed here
      }
      else {
        var rawResponse = fetcher(user, fullURL);
        response = rawResponse; // this particular return needs to be parsed here
      }
      break;
    case ART_ENUMS.components:
      fullURL += "/components?active_only=false&include_deleted=false&start_row=" + startRow + "&number_rows=" + numberOfRows + "&sort_field=ComponentId&sort_direction=ASC&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.customLists:
      if (artifactId > 0) {
        //if we have a valid custom list ID
        fullURL = API_TEMPLATE_BASE + templateId + "/custom-lists/" + artifactId + "?start_row=" + startRow + "&number_rows=" + numberOfRows + "&sort_field=CustomPropertyListId&sort_direction=ASC&";
      }
      else {
        //if not, retrieve all the lists
        fullURL = API_TEMPLATE_BASE + templateId + "/custom-lists?start_row=" + startRow + "&number_rows=" + numberOfRows + "&sort_field=CustomPropertyListId&sort_direction=ASC&";
      }
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.customValues:
      if (artifactId) {
        fullURL = API_TEMPLATE_BASE + templateId + "/custom-lists/" + artifactId + "?";
        response = fetcher(user, fullURL);
      }
      break;
  }
  return response;
}



/*
 *
 * =======================
 * CREATE "POST" FUNCTIONS
 * =======================
 *
 * functions to create new records in Spira - eg add new requirements
 *
 */

// General fetch function, using Google's built in fetch api
// @param: body - json object
// @param: currentUser - user object storing login data from client
// @param: postUrl - url string passed in to connect with Spira
function poster(body, currentUser, postUrl) {

  //use google's Utilities to base64 decode if present, otherwise use standard JS (ie for MS Excel)
  var decoded = typeof Utilities != "undefined" ? Utilities.base64Decode(currentUser.api_key) : atob(currentUser.api_key);
  var APIKEY = typeof Utilities != "undefined" ? Utilities.newBlob(decoded).getDataAsString() : decoded;

  //build URL from args
  var fullUrl = currentUser.url + postUrl + "username=" + currentUser.userName + APIKEY;

  //POST headers
  var params = {};
  params.method = 'post';
  params.contentType = 'application/json';
  params.muteHttpExceptions = true;
  if (body) params.payload = body;

  //call Google fetch function
  if (IS_GOOGLE) {
    var response = UrlFetchApp.fetch(fullUrl, params);
    return response;
  } else {
    //for MS Excel, use superagent to return a promise to the taskpane
    return superagent
      .post(fullUrl)
      .send(body)
      .set("Content-Type", "application/json", "accepts", "application/json")

  }
}



// General fetch function, using Google's built in fetch api
// @param: body - json object
// @param: currentUser - user object storing login data from client
// @param: PUTUrl - url string passed in to connect with Spira
function putUpdater(body, currentUser, PUTUrl) {
  //use google's Utilities to base64 decode if present, otherwise use standard JS (ie for MS Excel)
  var decoded = typeof Utilities != "undefined" ? Utilities.base64Decode(currentUser.api_key) : atob(currentUser.api_key);
  var APIKEY = typeof Utilities != "undefined" ? Utilities.newBlob(decoded).getDataAsString() : decoded;

  //build URL from args
  var fullUrl = currentUser.url + PUTUrl + "username=" + currentUser.userName + APIKEY;

  //PUT headers
  var params = {};
  params.method = 'put';
  params.contentType = 'application/json';
  params.muteHttpExceptions = true;
  if (body) params.payload = body;

  //call Google fetch function
  if (IS_GOOGLE) {
    var response = UrlFetchApp.fetch(fullUrl, params);
    return response;
  } else {
    //for MS Excel, use superagent to return a promise to the taskpane
    var putResult =
      superagent
        .put(fullUrl)
        .send(body)
        .set("Content-Type", "application/json", "accepts", "application/json");

    return putResult;
  }
}



// decides what is the association type based on the artifact type and entry
// returns the enum of the association type
function getAssociationType(artifactTypeId, entry) {

  //Associations for Test Cases:
  if (artifactTypeId == ART_ENUMS.testCases) {
    //TC and Requirement
    if (entry.RequirementId && entry.TestCaseId) {
      return params.associationEnums.tc2req;
    }
    //TC and Releases
    if (entry.ReleaseId && entry.TestCaseId) {
      return params.associationEnums.tc2rel;
    }
    //TC and TS
    if (entry.TestSetTestCaseId && entry.TestSetId) {
      return params.associationEnums.tc2ts;
    }
  }
  else if (artifactTypeId == ART_ENUMS.requirements) {
    return params.associationEnums.req2req;
  }
}



// effectively a switch to manage which comment we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactTypeId - int of the current artifact
function postCommentToSpira(entry, user, projectId, artifactTypeId) {
  //stringify
  var JSON_body = JSON.stringify(entry),
    response = "",
    postUrl = "";

  //send JSON object of new item to artifact specific export function
  switch (artifactTypeId) {

    // REQUIREMENTS
    case ART_ENUMS.requirements:
      postUrl = API_PROJECT_BASE + projectId + '/requirements/' + entry.ArtifactId + "/comments?";
      break;

    // TEST CASES
    case ART_ENUMS.testCases:
      postUrl = API_PROJECT_BASE + projectId + '/test-cases/' + entry.ArtifactId + "/comments?";
      break;

    // INCIDENTS
    case ART_ENUMS.incidents:
      postUrl = API_PROJECT_BASE + projectId + '/incidents/' + entry.ArtifactId + "/comments?";
      //creating comments for incidents requires the POST body to be array-like
      JSON_body = "[" + JSON_body + "]";
      break;

    // RELEASES
    case ART_ENUMS.releases:
      postUrl = API_PROJECT_BASE + projectId + '/releases/' + entry.ArtifactId + "/comments?";
      break;

    // TASKS
    case ART_ENUMS.tasks:
      postUrl = API_PROJECT_BASE + projectId + '/tasks/' + entry.ArtifactId + "/comments?";
      break;

    // RISKS
    case ART_ENUMS.risks:
      postUrl = API_PROJECT_BASE + projectId + '/risks/' + entry.ArtifactId + "/comments?";
      break;

    // TEST SETS
    case ART_ENUMS.testSets:
      postUrl = API_PROJECT_BASE + projectId + '/test-sets/' + entry.ArtifactId + "/comments?";
      break;
  }

  return postUrl ? poster(JSON_body, user, postUrl) : null;
}



// effectively a switch to manage which association we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactTypeId - int of the current artifact
function postAssociationToSpira(entry, user, projectId, artifactTypeId) {
  //stringify
  var JSON_body = JSON.stringify(entry),
    response = "",
    postUrl = "";

  //send JSON object of new item to artifact specific export function
  switch (artifactTypeId) {

    // REQUIREMENTS
    case ART_ENUMS.requirements:
      var associationType = getAssociationType(artifactTypeId, entry);
      if (associationType == params.associationEnums.req2req) {
        postUrl = API_PROJECT_BASE + projectId + '/associations?';
      }
      break;

    // TEST CASES
    case ART_ENUMS.testCases:
      //decide what association to handle
      var associationType = getAssociationType(artifactTypeId, entry);
      if (associationType == params.associationEnums.tc2req) {
        postUrl = API_PROJECT_BASE + projectId + '/requirements/test-cases?';
      }
      if (associationType == params.associationEnums.tc2rel) {
        postUrl = API_PROJECT_BASE + projectId + '/releases/' + entry.ReleaseId + '/test-cases?';
        //we need to handle this request in a special way
        JSON_body = '[' + entry.TestCaseId + ']';
      }
      if (associationType == params.associationEnums.tc2ts) {
        postUrl = API_PROJECT_BASE + projectId + '/test-sets/' + entry.TestSetId + '/test-case-mapping/' + entry.TestSetTestCaseId + '?';
        //we need to handle this request in a special way
        JSON_body = '';
      }
      break;
  }
  return postUrl ? poster(JSON_body, user, postUrl) : null;
}






// effectively a switch to manage which artifact we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactId - int of the current artifact
// @param: parentId - optional int of the relevant parent to attach the artifact too
function postArtifactToSpira(entry, user, projectId, templateId, artifactTypeId, parentId) {
  //stringify
  var JSON_body = JSON.stringify(entry),
    response = "",
    postUrl = "";
  //send JSON object of new item to artifact specific export function
  switch (artifactTypeId) {
    // REQUIREMENTS
    case ART_ENUMS.requirements:
      // url to post initial RQ to ensure it is fully outdented
      if (entry.indentPosition === 0) {
        postUrl = API_PROJECT_BASE + projectId + '/requirements/indent/' + INITIAL_HIERARCHY_OUTDENT + '?';
        // if no parentId then post as a regular RQ 
      } else if (parentId === -1) {
        postUrl = API_PROJECT_BASE + projectId + '/requirements?';
        // we should have a parent Id set so add this RQ as its child
      } else {
        postUrl = API_PROJECT_BASE + projectId + '/requirements/parent/' + parentId + '?';
      }
      break;

    // TEST CASES
    case ART_ENUMS.testCases:
      postUrl = API_PROJECT_BASE + projectId + '/test-cases?';
      break;

    // INCIDENTS
    case ART_ENUMS.incidents:
      postUrl = API_PROJECT_BASE + projectId + '/incidents?';
      break;

    // RELEASES
    case ART_ENUMS.releases:
      // if no parentId then post as a regular release
      if (parentId === -1) {
        postUrl = API_PROJECT_BASE + projectId + '/releases?';
        // we should have a parent Id set so add this RQ as its child
      } else {
        postUrl = API_PROJECT_BASE + projectId + '/releases/' + parentId + '?';
      }
      break;

    // TASKS
    case ART_ENUMS.tasks:
      postUrl = API_PROJECT_BASE + projectId + '/tasks?';
      break;

    // TEST STEPS
    case ART_ENUMS.testSteps:
      postUrl = parentId !== -1 ? API_PROJECT_BASE + projectId + '/test-cases/' + parentId + '/test-steps?' : null;
      // only post the test step if we have a parent id
      break;

    // RISKS
    case ART_ENUMS.risks:
      postUrl = API_PROJECT_BASE + projectId + '/risks?';
      entry['CreationDate'] = new Date().toISOString();
      JSON_body = JSON.stringify(entry);
      break;

    // TEST SETS
    case ART_ENUMS.testSets:
      postUrl = API_PROJECT_BASE + projectId + '/test-sets?';
      break;

    // FOLDERS
    case ART_ENUMS.folders:
      if (entry.artifact == ART_ENUMS.testCases) {
        postUrl = API_PROJECT_BASE + projectId + '/test-folders?';
      } else if (entry.artifact == ART_ENUMS.testSets) {
        postUrl = API_PROJECT_BASE + projectId + '/test-set-folders?';
      } else if (entry.artifact == ART_ENUMS.tasks) {
        postUrl = API_PROJECT_BASE + projectId + '/task-folders?';
      }
      break;

    // COMPONENTS
    case ART_ENUMS.components:
      postUrl = API_PROJECT_BASE + projectId + '/components?';
      break;

    // USERS
    case ART_ENUMS.users:

      postUrl = API_BASE + 'users?';
      //adding the optional URL paramenters
      if (entry.password) { postUrl += "password=" + entry.password + "&" };
      if (entry.password_question) { postUrl += "password_question=" + entry.password_question + "&" };
      if (entry.password_answer) { postUrl += "password_answer=" + entry.password_answer + "&" };
      if (entry.project_id) { postUrl += "project_id=" + entry.project_id + "&" };
      if (entry.project_role_id) { postUrl += "project_role_id=" + entry.project_role_id + "&" };

      //since the user is being created by an administrator, it should be already approved
      entry.Approved = "true";
      JSON_body = JSON.stringify(entry);
      break;

    //CUSTOM LISTS
    case ART_ENUMS.customLists:
      postUrl = API_TEMPLATE_BASE + templateId + '/custom-lists?';
      break;

    //CUSTOM VALUES
    case ART_ENUMS.customValues:
      postUrl = API_TEMPLATE_BASE + templateId + '/custom-lists/' + parentId + '/values?';
      entry[ART_PARENT_IDS[ART_ENUMS.customValues]] = parentId;
      JSON_body = JSON.stringify(entry);
      break;

  }
  return postUrl ? poster(JSON_body, user, postUrl) : null;
}



// effectively a switch to manage which artifact we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactId - int of the current artifact
// @param: parentId - optional int of the relevant parent to attach the artifact too
function putArtifactToSpira(entry, user, projectId, templateId, artifactTypeId, parentId) {
  var JSON_body = JSON.stringify(entry),
    response = "",
    putUrl = "";
  //send JSON object of new item to artifact specific export function
  switch (artifactTypeId) {

    // REQUIREMENTS
    case ART_ENUMS.requirements:
      // url to update RQs
      putUrl = API_PROJECT_BASE + projectId + '/requirements?';
      break;

    // TEST CASES
    case ART_ENUMS.testCases:
      putUrl = API_PROJECT_BASE + projectId + '/test-cases?';
      if (!IS_GOOGLE) {
        entry.TestSteps = null;
        JSON_body = JSON.stringify(entry);
      }
      break;

    // INCIDENTS
    case ART_ENUMS.incidents:
      putUrl = API_PROJECT_BASE + projectId + '/incidents/' + entry.IncidentId + '?';
      break;

    // RELEASES
    case ART_ENUMS.releases:
      // if no parentId then post as a regular release
      if (parentId === -1) {
        putUrl = API_PROJECT_BASE + projectId + '/releases?';
        // we should have a parent Id set so add this RQ as its child
      } else {
        putUrl = API_PROJECT_BASE + projectId + '/releases?';
      }
      break;

    // TASKS
    case ART_ENUMS.tasks:
      putUrl = API_PROJECT_BASE + projectId + '/tasks?';
      break;

    // TEST STEPS
    case ART_ENUMS.testSteps:
      putUrl = parentId !== -1 ? API_PROJECT_BASE + projectId + '/test-cases/' + parentId + '/test-steps?' : null;
      // only post the test step if we have a parent id
      break;

    // RISKS
    case ART_ENUMS.risks:
      putUrl = API_PROJECT_BASE + projectId + '/risks?';
      break;
    // TEST SETS
    case ART_ENUMS.testSets:
      putUrl = API_PROJECT_BASE + projectId + '/test-sets/?';
      break;
    //COMPONENTS
    case ART_ENUMS.components:
      putUrl = API_PROJECT_BASE + projectId + '/components/' + entry.ComponentId + '?';
      break;
    //CUSTOM LISTS
    case ART_ENUMS.customLists:
      putUrl = API_TEMPLATE_BASE + templateId + '/custom-lists/' + entry[ART_PARENT_IDS[ART_ENUMS.customLists]] + '?';
      break;
    //CUSTOM VALUES
    case ART_ENUMS.customValues:
      putUrl = API_TEMPLATE_BASE + templateId + '/custom-lists/' + entry[ART_PARENT_IDS[ART_ENUMS.customLists]] + '?';
      //due to the API nature, this query requires some extra steps before sending the object
      var finalEntry = {
        Name: "",
        Active: true,
        Values: []
      };
      finalEntry.Name = entry.parentName;
      finalEntry.Active = entry.parentActive;
      finalEntry.Values.push(entry);
      JSON_body = JSON.stringify(finalEntry);
      break;
  }

  return putUrl ? putUpdater(JSON_body, user, putUrl) : null;
}

/*
 *
 * ==============
 * ERROR MESSAGES
 * ==============
 *
 */

// Error notification function
// Assigns string value and routes error call from client.js.html
// @param: type - string identifying the message to be displayed
// @param: err - the detailed error object (differs between plugin)
function error(type, err) {
  var message = "",
    details = "";
  if (type == 'impExp') {
    message = 'There was an input error. Please check that your entries are correct.';
  } else if (type == "network") {
    message = 'Network error. Please check your username, url, and password. If correct make sure you have the correct permissions.';
    details = err ? `<br><br><b>STATUS:</b> ${err.status ? err.status : "unknown"}<br><br><b>MESSAGE:</b> ${err.stack ? err.stack : "unknown"}` : "";
  } else if (type == 'excel') {
    message = 'Excel reported an error!';
    details = err ? `<br><br>Description: ${err.description}` : "";
  } else if (type == 'unknown' || err == 'unknown') {
    message = 'Unknown error. Please try again later or contact your system administrator';
  } else if (type == 'sheet') {
    message = 'There was a problem while retrieving data from the active spreadsheet. Please check the details below and try again. <br><br><b>Details:</b><br>' + err;
  } else {
    message = 'Unknown error. Please try again later or contact your system administrator';
  }

  if (IS_GOOGLE) {
    okWarn(message);
  } else {
    popupShow(message + details, "");
  }
}



// Pop-up notification function
// @param: string - message to be displayed
function success(string) {
  // Show a 2-second popup with the title "Status" and a message passed in as an argument.
  SpreadsheetApp.getActiveSpreadsheet().toast(string, 'Success', 2);
}



// Alert pop up for data clear warning
// @param: string - message to be displayed
function warn(string) {
  var ui = SpreadsheetApp.getUi();
  //alert popup with yes and no button
  var response = ui.alert(string, ui.ButtonSet.YES_NO);

  //returns with user choice
  if (response == ui.Button.YES) {
    return true;
  } else {
    return false;
  }
}



// Alert pop up for export success
// @param: message - string sent from the export function
// @param: isTemplateLoadFail - bool about whether this message means that the template load sequence has failed
function operationComplete(messageEnum, isTemplateLoadFail) {
  if (IS_GOOGLE) {
    var message = STATUS_MESSAGE_GOOGLE[messageEnum] || STATUS_MESSAGE_GOOGLE['1'];
    okWarn(message);
  } else {
    var message = STATUS_MESSAGE_EXCEL[messageEnum] || STATUS_MESSAGE_EXCEL['1'];
    return popupShow(message, "", isTemplateLoadFail);
  }
}

// Alert pop up for no template present
function noTemplate() {
  okWarn('Please load a template to continue.');
}



// Google alert popup with OK button
// @param: dialog - message to show
function okWarn(dialog) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(dialog, ui.ButtonSet.OK);
}


/*
 * =================
 * TEMPLATE CREATION
 * =================
 *
 * This function creates a template based on the model template data
 * Takes the entire data model as an argument
 *
 */

// function that manages template creation - creating the header row, formatting cells, setting validation
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldTypeEnums - list of fieldType enums from client params object
function templateLoader(model, fieldTypeEnums, advancedMode) {

  var fields = model.fields;
  var sheet;
  var newSheetName;

  if (model.currentOperation) {
    //administrator mode

    var operation = model.operations.filter(function (operation) {
      return operation.id == model.currentOperation;
    })[0];

    if (operation.type == "send-system") {
      //system wide operations
      newSheetName = model.currentArtifact.name + ", system";

    }
    else if (operation.type == "send-template" || operation.type == "get-template") {
      //template-based operations
      newSheetName = model.currentArtifact.name + ", TP-" + model.currentTemplate.id;
    }
  }
  else {
    //stardard artifact functions
    newSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
  }

  var response;
  // select active sheet
  if (IS_GOOGLE) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // set sheet (tab) name to model name
    sheet.setName(newSheetName);
    sheetSetForTemplate(sheet, model, fieldTypeEnums, null);
    response = true;
  } else {

    if (!advancedMode) {
      //if not in advanced mode, ignore the fields only available for that mode

      model.fields = fields.filter(function (item, index) {
        if (!item.isAdvanced) {
          return item;
        }
      })
    }

    return Excel.run(function (context) {
      // store the sheet and worksheet list for use later
      sheet = context.workbook.worksheets.getActiveWorksheet();

      //reset the hidden status of the spreadsheet
      var range = sheet.getRangeByIndexes(0, 0, 1, EXCEL_MAX_ROWS_PER_PAGE);
      range.columnHidden = false;

      var worksheets = context.workbook.worksheets;
      worksheets.load('items');

      return context.sync()
        .then(function () {
          let onWrongSheet = false;

          // check that no other worksheet has the same name as the one we need to call this sheet
          if (worksheets.items.length > 1) {
            worksheets.items.forEach(x => {
              if (x.name == newSheetName && x.id !== sheet.id) {
                onWrongSheet = true;
              }
            });
          }

          if (onWrongSheet) {
            return operationComplete(STATUS_ENUM.wrongSheet, true);
          } else {
            // otherwise set the sheet name, then create the template
            sheet.name = newSheetName;
            return context.sync()
              .then(function () {
                return sheetSetForTemplate(sheet, model, fieldTypeEnums, context, newSheetName);
              })
          }
        })
        .catch(/*fail quietly*/);
    })
  }
}

// wrapper function to set the header row, validation rules, and any extra formatting 
function sheetSetForTemplate(sheet, model, fieldTypeEnums, context, newSheetName) {

  // heading row - sets names and formatting (standard sheet) 
  headerSetter(sheet, model.fields, model.colors, context);
  // set validation rules on the columns (standard sheet) 
  contentValidationSetter(sheet, model, fieldTypeEnums, context);
  // set any extra formatting options (standard sheet) 
  contentFormattingSetter(sheet, model, context);
  //set database fields (database sheet) 
  dataBaseValidationSetter(newSheetName, model, fieldTypeEnums, context);
}



// Sets headings for fields
// creates an array of the field names so that changes can be batched to the relevant range in one go for performance reasons
// @param: sheet - the sheet object
// @param: fields - full field data
// @param: colors - global colors used for formatting
function headerSetter(sheet, fields, colors, context) {

  var headerNames = [],
    backgrounds = [],
    fontColors = [],
    fontWeights = [],
    fieldsLength = fields.length;

  for (var i = 0; i < fieldsLength; i++) {
    headerNames.push(fields[i].name);

    // set field text depending on whether is required or not
    var fontColor = (fields[i].required || fields[i].requiredForSubType) ? colors.headerRequired : colors.header;
    var fontWeight = fields[i].required ? 'bold' : 'normal';
    fontColors.push(fontColor);
    fontWeights.push(fontWeight);

    // set background colors based on if it is a subtype only field or not
    var background = fields[i].isSubTypeField ? colors.bgHeaderSubType : colors.bgHeader;
    //check for special dual type fields
    if (fields[i].isTypeAndSubTypeField) { background = colors.bgHeaderTypeAndSubType; }
    backgrounds.push(background);
  }

  if (IS_GOOGLE) {
    sheet.getRange(1, 1, 1, fieldsLength)
      .setWrap(true)
      // the arrays need to be in an array as methods expect a 2D array for managing 2D ranges
      .setBackgrounds([backgrounds])
      .setFontColors([fontColors])
      .setFontWeights([fontWeights])
      .setValues([headerNames])
      .protect().setDescription("header row").setWarningOnly(true);

  } else {
    var range = sheet.getRangeByIndexes(0, 0, 1, fieldsLength);
    range.values = [headerNames];
    for (var i = 0; i < fieldsLength; i++) {
      var cellRange = sheet.getCell(0, i);
      cellRange.set({
        format: {
          fill: { color: backgrounds[i] },
          font: { color: fontColors[i], bold: fontWeights[i] == "bold" }
        }
      });
    }
    return context.sync();
  }
}



// Sets validation on a per column basis, based on the field type passed in by the model
// a switch statement checks for any type requiring validation and carries out necessary action
// @param: sheet - the sheet object
// @param: model - full data to acccess global params as well as all fields
// @param: fieldTypeEnums - enums for field types
function contentValidationSetter(sheet, model, fieldTypeEnums, context) {
  // we can't easily get the max rows for excel so use the number of rows it always seems to have
  var nonHeaderRows = IS_GOOGLE ? sheet.getMaxRows() - 1 : EXCEL_NUMBER_OF_ROWS - 1;

  for (var index = 0; index < model.fields.length; index++) {
    var columnNumber = index + 1,
      list = [];

    switch (model.fields[index].type) {

      // ID fields: restricted to numbers and protected
      case fieldTypeEnums.id:
      case fieldTypeEnums.subId:
        setPositiveIntValidation(sheet, columnNumber, nonHeaderRows, false);
        protectColumn(
          sheet,
          columnNumber,
          nonHeaderRows,
          model.colors.bgReadOnly,
          "ID field"
        );
        break;

      // INT and NUM fields are both treated by Sheets as numbers
      case fieldTypeEnums.int:
        setIntegerValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customInteger:
        setIntegerValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customAutomationHost:
        setIntegerValidation(sheet, columnNumber, nonHeaderRows, false);
        break;


      // NUM fields are handled as decimals by Excel though
      case fieldTypeEnums.num:
        setNumberValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customDecimal:
        setNumberValidation(sheet, columnNumber, nonHeaderRows, false);
        break;

      // DATE fields get date validation
      case fieldTypeEnums.date:
        setDateValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customDate:
        setDateValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customDateAndTime:
        setDateTimeValidation(sheet, columnNumber, nonHeaderRows, false);
        break;

      //TEXT types validation
      case fieldTypeEnums.text:
        setTextValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customText:
        setTextValidation(sheet, columnNumber, nonHeaderRows, false);
        break;
      case fieldTypeEnums.customPassword:
        setTextValidation(sheet, columnNumber, nonHeaderRows, false);
        break;

      // All other types
      default:
        //do nothing
        break;
    }
  }
}

//Function used to set the databaseSheet complete name, given the project and artifact ID
function createDatabaseSheetName(sheetName, projectId, artifactTypeId) {
  return sheetName + "-" + projectId + "-" + artifactTypeId;
}

// Sets validation on a per column basis, based on the field type passed in by the model
// a switch statement checks for any type requiring validation and carries out necessary action
// @param: sheet - the sheet object
// @param: model - full data to acccess global params as well as all fields
// @param: fieldTypeEnums - enums for field types
function dataBaseValidationSetter(mainSheetName, model, fieldTypeEnums, context) {
  // we can't easily get the max rows for excel so use the number of rows it always seems to have
  for (var index = 0; index < model.fields.length; index++) {
    var columnNumber = index + 1,
      list = [];
    switch (model.fields[index].type) {
      // DROPDOWNS and MULTIDROPDOWNS are both treated as simple dropdowns (Sheets does not have multi selects)
      case fieldTypeEnums.drop:
      case fieldTypeEnums.customList:
      case fieldTypeEnums.multi:
      case fieldTypeEnums.customMultiList:
        var fieldList = model.fields[index].values;
        for (var i = 0; i < fieldList.length; i++) {
          list.push(setListItemDisplayName(fieldList[i]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
        break;

      // RELEASE fields are dropdowns with the values coming from a project wide set list
      case fieldTypeEnums.release:
        if (!model.fields[index].showInactiveReleases) {
          //show only the active releases
          for (var l = 0; l < model.projectActiveReleases.length; l++) {
            list.push(setListItemDisplayName(model.projectActiveReleases[l]));
          }
          setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
        }
        else {
          //show all the releases - including inactives
          for (var l = 0; l < model.projectReleases.length; l++) {
            list.push(setListItemDisplayName(model.projectReleases[l]));
          }
          setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
          break;
        }
        break;
      case fieldTypeEnums.customRelease:
        for (var l = 0; l < model.projectReleases.length; l++) {
          list.push(setListItemDisplayName(model.projectReleases[l]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
        break;

      // BOOL as Sheets has no bool validation, a yes/no dropdown is used
      case fieldTypeEnums.bool:
        // 'True' and 'False' don't work as dropdown choices
        list.push("Yes", "No");
        setDropdownValidation(mainSheetName, columnNumber, list, false, context, model);
        break;

      case fieldTypeEnums.customBoolean:
        // 'True' and 'False' don't work as dropdown choices
        list.push("Yes", "No");
        setDropdownValidation(mainSheetName, columnNumber, list, false, context, model);
        break;


      // USER fields are dropdowns with the values coming from a project wide set list
      case fieldTypeEnums.user:
        for (var j = 0; j < model.projectUsers.length; j++) {
          list.push(setListItemDisplayName(model.projectUsers[j]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
        break;

      case fieldTypeEnums.customUser:
        for (var j = 0; j < model.projectUsers.length; j++) {
          list.push(setListItemDisplayName(model.projectUsers[j]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
        break;

      // COMPONENT fields are dropdowns with the values coming from a project wide set list
      case fieldTypeEnums.component:
        for (var k = 0; k < model.projectComponents.length; k++) {
          list.push(setListItemDisplayName(model.projectComponents[k]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, true, context, model);
        break;

      // All other types
      default:
        //do nothing
        break;
    }
  }
  if (!IS_GOOGLE) {
    return context.sync();
  }
}


// create dropdown validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: list - array of values to show in a dropdown and use for validation
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
async function setDropdownValidation(mainSheetName, columnNumber, list, allowInvalid, context, model) {
  if (IS_GOOGLE) {
    //first, write the values to the dbSheet
    var dataBaseSheetName = createDatabaseSheetName(params.dataSheetName, model.currentProject.id, model.currentArtifact.id);

    if (model.currentOperation) {
      //administrator mode

      var operation = model.operations.filter(function (operation) {
        return operation.id == model.currentOperation;
      })[0];

      if (operation.type == "send-system") {
        //system wide operations
        mainSheetName = model.currentArtifact.name + ", system";

      }
      else if (operation.type == "send-template" || operation.type == "get-template") {
        //template-based operations
        mainSheetName = model.currentArtifact.name + ", TP-" + model.currentTemplate.id;
      }
    }
    else {
      //stardard artifact functions
      mainSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
    }

    var values = [];
    list.forEach(function (item) {
      var itemArray = [item];
      values.push(itemArray);
    });

    var dbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataBaseSheetName),
      dbRange = dbSheet.getRange(1, columnNumber, list.length, 1);
    dbRange.setValues(values);

    //Now, point the fields in the mainsheet to the database worksheet (source)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetName);
    var range = sheet.getRange(2, columnNumber, sheet.getMaxRows() - 1);
    // create the validation rule
    // requireValueInList - params are the array to use, and whether to create a dropdown list
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(dbRange, true)
      .setAllowInvalid(allowInvalid)
      .build();
    range.setDataValidation(rule);
  } else {
    //max rows for Excel
    var nonHeaderRows = EXCEL_NUMBER_OF_ROWS - 1;
    //first, write the values to the dbSheet
    var values = [];
    list.forEach(function (item) {
      var itemArray = [item];
      values.push(itemArray);
    });

    var dataBaseSheetName = createDatabaseSheetName(params.dataSheetName, model.currentProject.id, model.currentArtifact.id);

    var dbSheetRange = context.workbook.worksheets.getItem(dataBaseSheetName).getRangeByIndexes(0, columnNumber - 1, list.length, 1);
    dbSheetRange.values = values;
    context.sync();
    //Now, point the fields in the mainsheet to the database worksheet (source)
    var range = context.workbook.worksheets.getItem(mainSheetName).getRangeByIndexes(1, columnNumber - 1, nonHeaderRows, 1);
    range.dataValidation.clear();

    range.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source: dbSheetRange
      }
    };
    await context.sync();
  }
}



// create date validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setDateValidation(sheet, columnNumber, rowLength, allowInvalid) {
  if (IS_GOOGLE) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    var rule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText('Must be a valid date')
      .build();
    range.setDataValidation(rule);
    range.setNumberFormat('mm/dd/yy');

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();


    var greaterThan2000Rule = {
      date: {
        formula1: "2000-01-01",
        operator: Excel.DataValidationOperator.greaterThan
      }
    };
    range.dataValidation.rule = greaterThan2000Rule;

    range.dataValidation.prompt = {
      message: "Please enter a date.",
      showPrompt: true,
      title: "Valid dates only."
    };
    range.dataValidation.errorAlert = {
      message: "Sorry, only dates are allowed",
      showAlert: true,
      style: "Stop",
      title: "Invalid date entered"
    };
    //now set the cell format to dates
    range.numberFormatLocal = "dd-mmm-yyyy";
  }
}

// create dateTime validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setDateTimeValidation(sheet, columnNumber, rowLength, allowInvalid) {
  if (IS_GOOGLE) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    var rule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText('Must be a valid date')
      .build();
    range.setDataValidation(rule);
    range.setNumberFormat('mm/dd/yy hh:mm:ss am/pm');

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    var greaterThan2000Rule = {
      date: {
        formula1: "2000-01-01",
        operator: Excel.DataValidationOperator.greaterThan
      }
    };
    range.dataValidation.rule = greaterThan2000Rule;

    range.dataValidation.prompt = {
      message: "Please enter a date.",
      showPrompt: true,
      title: "Valid dates only."
    };
    range.dataValidation.errorAlert = {
      message: "Sorry, only dates are allowed",
      showAlert: true,
      style: "Stop",
      title: "Invalid date entered"
    };
    //now set the cell format to dates
    range.numberFormatLocal = "dd-mmm-yyyy hh:mm:ss am/pm";
  }
}



// create positive integer validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setPositiveIntValidation(sheet, columnNumber, rowLength, allowInvalid) {
  // create range
  if (IS_GOOGLE) {
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    //must be a valid number greater than -1 (also excludes 1.1.0 style numbers)
    var rule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(-1)
      .setAllowInvalid(true)
      //.setHelpText('Must be a positive number')
      .build();
    range.setDataValidation(rule);

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    var greaterThanZeroRule = {
      wholeNumber: {
        formula1: 0,
        operator: Excel.DataValidationOperator.greaterThan
      }
    };
    range.dataValidation.rule = greaterThanZeroRule;

    range.dataValidation.prompt = {
      message: "Please enter a positive number.",
      showPrompt: true,
      title: "Positive numbers only."
    };
    range.dataValidation.errorAlert = {
      message: "Sorry, only positive numbers are allowed",
      showAlert: true,
      style: "Stop",
      title: "Negative Number Entered"
    };
  }
}

// create number validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setIntegerValidation(sheet, columnNumber, rowLength, allowInvalid) {
  // create range
  if (IS_GOOGLE) {
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    //must be a valid number greater than -1 (also excludes 1.1.0 style numbers)
    var rule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(allowInvalid)
      .setHelpText('Must be a whole number!')
      .build();
    range.setDataValidation(rule);
    range.setNumberFormat('0');

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    var greaterThanZeroRule = {
      wholeNumber: {
        formula1: 0,
        operator: Excel.DataValidationOperator.greaterThan
      }
    };
    range.dataValidation.rule = greaterThanZeroRule;

    range.dataValidation.prompt = {
      message: "Please enter a valid number.",
      showPrompt: true,
      title: "Whole Numbers only"
    };
    range.dataValidation.errorAlert = {
      message: "Sorry, only whole numbers are allowed.",
      showAlert: true,
      style: "Stop",
      title: "Invalid entry"
    };
  }
}

// create decimal validation on set column based on specified values - identical to integer validation for Google Sheets as of 2020-07
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setNumberValidation(sheet, columnNumber, rowLength, allowInvalid) {
  // create range
  if (IS_GOOGLE) {
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    //must be a valid number greater than -1 (also excludes 1.1.0 style numbers)
    var rule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(allowInvalid)
      .setHelpText('Must be a number!')
      .build();
    range.setDataValidation(rule);
    range.setNumberFormat('####.##');

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    var greaterThanZeroRule = {
      decimal: {
        formula1: 0,
        operator: Excel.DataValidationOperator.greaterThan
      }
    };
    range.dataValidation.rule = greaterThanZeroRule;

    range.dataValidation.prompt = {
      message: "Please enter a valid number.",
      showPrompt: true,
      title: "Numbers only"
    };
    range.dataValidation.errorAlert = {
      message: "Sorry, only numbers are allowed.",
      showAlert: true,
      style: "Stop",
      title: "Invalid entry"
    };
  }
}

// create text validation on set column base
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setTextValidation(sheet, columnNumber, rowLength, allowInvalid) {
  if (IS_GOOGLE) {
    var range = sheet.getRange(2, columnNumber, rowLength);
    // create the validation rule
    var rule = SpreadsheetApp.newDataValidation()
      .setAllowInvalid(allowInvalid)
      .requireTextContains('')
      .setHelpText('Must be a text string')
      .build();
    range.setDataValidation(rule);
    range.setNumberFormat('@STRING@');

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    range.numberFormat = '@';

    range.dataValidation.errorAlert = {
      message: "Sorry, only text string is allowed.",
      showAlert: true,
      style: "Stop",
      title: "Invalid entry"
    };
  }
}

// format columns based on a potential range of factors - eg hide unsupported columns
// @param: sheet - the sheet object
// @param: model - full model data set
function contentFormattingSetter(sheet, model) {
  for (var i = 0; i < model.fields.length; i++) {
    var columnNumber = i + 1;
    var nonHeaderRows = IS_GOOGLE ? sheet.getMaxRows() - 1 : EXCEL_NUMBER_OF_ROWS - 1;

    // protect column
    // read only fields - ie ones you can get from Spira but not create in Spira (as with IDs - eg task component)
    if (model.fields[i].unsupported || model.fields[i].isReadOnly) {
      var warning = "";
      if (model.fields[i].unsupported) {
        warning = model.fields[i].name + "unsupported";
      } else if (model.fields[i].isReadOnly) {
        warning = model.fields[i].name + " is read only";
      }

      protectColumn(
        sheet,
        columnNumber,
        nonHeaderRows,
        model.colors.bgReadOnly,
        warning
      );
    }
    // hide this column if specified in the field model
    if (model.fields[i].isHidden) {
      var warning = "";
      warning = model.fields[i].name + " is hidden";
      hideColumn(
        sheet,
        columnNumber,
        nonHeaderRows,
        model.colors.bgReadOnly
      );

    }

  }
}



// protects specific column. Edits still allowed - current user not excluded from edit list, but could in future
// @param: sheet - the sheet object
// @param: columnNumber - int of column to hide
// @param: rowLength - int of default number of rows to apply any formattting to
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
// @param: name - string description for the protected range
function protectColumn(sheet, columnNumber, rowLength, bgColor, name) {
  // only for google as cannot protect individual cells easily in Excel
  if (IS_GOOGLE) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);
    range.setBackground(bgColor)
      .protect()
      .setDescription(name)
      .setWarningOnly(true);

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);

    // set the background color
    range.set({ format: { fill: { color: bgColor } } });

    // now we can add data validation
    // the easiest hack way to not allow any entry into the cell is to make sure its text length can only be zero
    range.dataValidation.clear();
    var textLengthZero = {
      textLength: {
        formula1: 0,
        operator: Excel.DataValidationOperator.equalTo
      }
    };
    range.dataValidation.rule = textLengthZero;

    range.dataValidation.prompt = {
      message: "This is a protected field and not user editable.",
      showPrompt: true,
      title: "No entry allowed."
    };
    range.dataValidation.errorAlert = {
      message: "Sorry, this is a protected field",
      showAlert: true,
      style: "Stop",
      title: "No entry allowed"
    };
  }

}


// hides a specific column range
// @param: sheet - the sheet object
// @param: columnNumber - int of column to hide
// @param: rowLength - int of default number of rows to apply any formattting to
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
function hideColumn(sheet, columnNumber, rowLength, bgColor) {
  // only for google as cannot protect individual cells easily in Excel
  if (IS_GOOGLE) {

    sheet.hideColumns(columnNumber);

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);

    // set the background color
    range.set({ format: { fill: { color: bgColor } } });

    //hide the column
    range.columnHidden = true;
  }
}

// resets backgroung colors to its original - used before a GET command
// @param: sheetName - current sheet name

function resetSheetColors(model, fieldTypeEnums, currentSheet) {

  if (IS_GOOGLE) {

    var fields = model.fields;
    var columnCount = Object.keys(fields).length + 1;
    var rowCount = currentSheet.getLastRow();
    if (rowCount == 0) { rowCount = 1; } //avoid errors

    //reset each column color schema (depending on property type)
    for (var j = 1; j < columnCount; j++) {
      var isHidden = fields[j - 1].isHidden;
      if (!isHidden) {
        var subColumnRange = currentSheet.getRange(2, j, rowCount, 1);
        var fieldType = fields[j - 1].type;
        var isReadOnly = fields[j - 1].isReadOnly;

        if (fieldType == fieldTypeEnums.id || fieldType == fieldTypeEnums.subId || isReadOnly) {
          subColumnRange.setBackground(model.colors.bgReadOnly);
          subColumnRange.clearNote();
        }
        else {
          subColumnRange.setBackground(model.colors.header);
          subColumnRange.clearNote();
        }
      }
    }
  }
  else {
    Excel.run(function (ctx) {
      var fields = model.fields;
      var columnCount = Object.keys(fields).length;

      //get the previous data number of rows
      var rowCount;
      var sheetOldData = sheetRangeOld.values;

      for (var i = 0; i < sheetOldData.length; i++) {
        // stop at the first row that is fully blank
        if (sheetOldData[i].join("") === "") {
          break;
        }
        else {
          rowCount = i;
        }
      }

      //complete data range from old data
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRangeByIndexes(1, 0, rowCount + 1, columnCount);

      //reset each column color schema (depending on property type)
      for (var j = 0; j < columnCount; j++) {

        var subColumnRange = range.getColumn(j);
        var fieldType = fields[j].type;
        var isReadOnly = fields[j].isReadOnly;
        var bgColor;

        if (fieldType == fieldTypeEnums.id || fieldType == fieldTypeEnums.subId || isReadOnly) {

          subColumnRange.format.fill.color = model.colors.bgReadOnly;
        }
        else {

          subColumnRange.format.fill.clear();
        }

      }
      range.delete(Excel.DeleteShiftDirection.up);
      return ctx.sync();
    }).catch(function (error) {
      if (error instanceof OfficeExtension.Error) {
        //fail quietly
      }
    });

  }
}



// removes error messages from the 1st field of each row, in case there's any. Only the IDs are preserved
// @param: sheetData - the sheet data object
// returns the filtered sheetData object

function clearErrorMessages(sheetData, isUpdate) {
  if (IS_GOOGLE) {
    // get active spreadsheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = spreadSheet.getActiveSheet(),
      lastColumn = sheet.getLastColumn(),
      lastRow = sheet.getLastRow();

    //minimum requirements
    if (lastColumn == 0) { lastColumn = 1 };
    if (lastRow == 0) { lastRow = 1 };

    // clears notes from the entire sheet
    var range = sheet.getRange(1, 1, lastRow, lastColumn);
    range.clearNote();

    for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
      //add -1 as ID for blank (invalid) new PUT values
      if (isUpdate && !sheetData[rowToPrep][0]) {
        sheetData[rowToPrep][0] = "-1";
      }
    }
  }
  else {
    for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
      // stop at the first row that is fully blank
      if (sheetData[rowToPrep].join("") === "") {
        break;
      } else {
        //the error messages are placed always at the position 0 of each row
        if (isNaN(sheetData[rowToPrep][0])) {
          try {
            sheetData[rowToPrep][0] = sheetData[rowToPrep][0].split(",")[0];
          }
          catch (err) {
            //do nothing
          }

        }
        //also, add -1 as ID for blank (invalid) new PUT values
        if (isUpdate && !sheetData[rowToPrep][0]) {
          sheetData[rowToPrep][0] = "-1";
        }
      }
    }
  }
  return sheetData;
}



// returns the artifact IDs containing validation problems
// @param: sheetData - the sheet data object
// returns the filtered sheetData object
function getValidationsFromEntries(entries, sheetData, response) {
  var artifactIds = [];
  entries.forEach(function isValidation(item, index) {
    if (index < response.entries.length) {

      if ('validationMessage' in item || 'error' in response.entries[index]) {
        //if this is a invalid entry, appends its ID to the array
        artifactIds.push(sheetData[index][0]);
      }
    }
  });
  return artifactIds;
}




/*
 * ================
 * SENDING TO SPIRA
 * ================
 *
 * The main function takes the entire data model and the artifact type
 * and calls the child function to set various object values before
 * sending the finished objects to SpiraPlan
 *
 */

// function that manages exporting data from the sheet - creating an array of objects based on entered data, then sending to Spira
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldTypeEnums - list of fieldType enums from client params object
// @param: isUpdate - boolean that indicates if this is an update operation (true) or create operation (false)
async function sendToSpira(model, fieldTypeEnums, isUpdate) {
  // 0. SETUP FUNCTION LEVEL VARS
  var entriesLog, commentsLog, associationsLog;

  var fields = model.fields,
    artifact = model.currentArtifact,
    artifactIsHierarchical = artifact.hierarchical;

  var requiredSheetName;

  if (model.currentOperation) {
    //administrator mode

    var operation = model.operations.filter(function (operation) {
      return operation.id == model.currentOperation;
    })[0];

    if (operation.type == "send-system") {
      //system wide operations
      requiredSheetName = model.currentArtifact.name + ", system";

    }
    else if (operation.type == "send-template" || operation.type == "get-template") {
      //template-based operations
      requiredSheetName = model.currentArtifact.name + ", TP-" + model.currentTemplate.id;
    }
  }
  else {
    //stardard artifact functions
    requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
  }

  // 1. get the active spreadsheet and first sheet
  if (IS_GOOGLE) {
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = spreadSheet.getActiveSheet(),
      lastRow = sheet.getLastRow() - 1 || 10, // hack to make sure we pass in some rows to the sheetRange, otherwise it causes an error
      sheetRange = sheet.getRange(2, 1, lastRow, fields.length),
      sheetData = sheetRange.getValues();

    if (sheet.getName() == requiredSheetName) {
      //Clear error messages from the ID fields, if any
      sheetData = clearErrorMessages(sheetData, isUpdate);
      //use this variable to save the new artifact entries
      var artifactEntries;
      //First, send the artifact entries for Spira
      var entriesForExport = createExportEntries(sheetData, model, fieldTypeEnums, fields, artifact, artifactIsHierarchical, isUpdate);
      var entriesLog = sendExportEntriesGoogle(entriesForExport, '', '', sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, '');
      var validations = getValidationsFromEntries(entriesForExport, sheetData, entriesLog);
      artifactEntries = entriesLog.entries;

      //Then, send the comments entry for Spira
      var commentEntriesForExport = createExportCommentEntries(sheetData, model, fieldTypeEnums, artifact, artifactIsHierarchical, artifactEntries, entriesLog);
      var commentsLog = sendExportEntriesGoogle('', commentEntriesForExport, '', sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, validations);

      //Finally, send the associations
      var associationEntriesForExport = createExportAssociationEntries(sheetData, model, fieldTypeEnums, artifact, artifactEntries, entriesLog);
      var associationsLog = sendExportEntriesGoogle('', '', associationEntriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, validations);
      var finalLog = updateSheetWithExportResults(entriesLog, commentsLog, associationsLog, entriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, null);
      return JSON.stringify(finalLog);
    } else {
      var errorLog = {
        status: STATUS_ENUM.wrongSheet
      };;
      return errorLog.toString();
    }
  } else {
    return await Excel.run({ delayForCellEdit: true }, function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet(),
        sheetRange = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS_PER_PAGE, fields.length);
      sheet.load("name");
      sheetRange.load("values");

      return new Promise(function (resolve, reject) {
        context.sync()
          .then(function () {
            if (sheet.name == requiredSheetName) {
              var sheetData = sheetRange.values;
              //Clear error messages from the ID fields, if any
              sheetData = clearErrorMessages(sheetData, isUpdate);
              //use this variable to save the new artifact entries
              var artifactEntries;
              var validations = [];
              //First, send the artifact entries for Spira
              var entriesForExport = createExportEntries(sheetData, model, fieldTypeEnums, fields, artifact, artifactIsHierarchical, isUpdate);
              return sendExportEntriesExcel(entriesForExport, '', '', sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context, isUpdate, '').then(function (response) {
                validations = getValidationsFromEntries(entriesForExport, sheetData, response);
                entriesLog = response;
                artifactEntries = response.entries;
                //Then, send the comments entry for Spira
                var commentEntriesForExport = createExportCommentEntries(sheetData, model, fieldTypeEnums, artifact, artifactIsHierarchical, artifactEntries, entriesLog);
                return sendExportEntriesExcel('', commentEntriesForExport, '', sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context, isUpdate, validations);
              }).then(function (responseComments) {
                commentsLog = responseComments;
                //Finally, send the associations
                var associationEntriesForExport = createExportAssociationEntries(sheetData, model, fieldTypeEnums, artifact, artifactEntries, entriesLog);
                return sendExportEntriesExcel('', '', associationEntriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context, isUpdate, validations);
              }).then(function (responseAssociations) {
                associationsLog = responseAssociations;
              }).catch(function (err) {
                reject()
              }).finally(function () {
                resolve(updateSheetWithExportResults(entriesLog, commentsLog, associationsLog, entriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, context))
              })
            }
            else {
              var log = {
                status: STATUS_ENUM.wrongSheet
              };
              return log;
            }
          })
          .catch();
      })
        .catch();
    })
  }
}



// function that verifies if a row correspond to a special case, that should be skipped
function isSpecialCase(rowToPrep, model, fieldTypeEnums, fields, artifact, artifactIsHierarchical, isUpdate) {
  // 1. Linked Test Steps
  var skipSpecial = false;
  for (const specialCase of params.specialCases) {

    //if true, we already found it
    var parameterIndex = 1;
    var fieldIndex = 4;

    //checking if it's a test case and it's a test step 
    if (specialCase.artifactId == model.currentArtifact.id && rowToPrep[parameterIndex]) {
      var stepName = rowToPrep[fieldIndex];
      if (stepName.startsWith(specialCase.target)) {
        skipSpecial = true;
        break;
      }
    }
  }
  if (skipSpecial) { return true; }
  else {
    return false;
  }

}

//function that verifies if a test Step has a valid TestCase parent
function isValidParent(entriesForExport) {
  for (let index = entriesForExport.length; index > 0; index--) {
    //if this is not a test step, we found the parent index
    if (!entriesForExport[index - 1].isSubType) {
      if (entriesForExport[index - 1].validationMessage) {
        //this is an invalid parent
        return false;
      }
      else {
        //this is a valid parent
        return true;
      }
    }
  }
  return true;
}



// 2. CREATE ARRAY OF ENTRIES
//2.1 Custom and Standard fields - Sent through the artifact API function (POST/PUT)
// loop to create artifact objects from each row taken from the spreadsheet
// vars needed: sheetData, artifact, fields, model, fieldTypeEnums, artifactIsHierarchical,
function createExportEntries(sheetData, model, fieldTypeEnums, fields, artifact, artifactIsHierarchical, isUpdate) {
  var lastIndentPosition = null,
    entriesForExport = [],
    isComment = false;

  for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
    // stop at the first row that is fully blank
    if (sheetData[rowToPrep].join("") === "") {
      break;
    } else {
      // check for required fields (for normal artifacts and those with sub types - eg test cases and steps)
      var skipSpecial = isSpecialCase(sheetData[rowToPrep], model, fieldTypeEnums, fields, artifact, artifactIsHierarchical, isUpdate);
      if (!skipSpecial) {
        var rowChecks = {
          hasSubType: artifact.hasSubType,
          totalFieldsRequired: countRequiredFieldsByType(fields, false),
          totalSubTypeFieldsRequired: artifact.hasSubType ? countRequiredFieldsByType(fields, true) : 0,
          countRequiredFields: rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, false),
          countSubTypeRequiredFields: artifact.hasSubType ? rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, true) : 0,
          subTypeIsBlocked: !artifact.hasSubType ? true : rowBlocksSubType(sheetData[rowToPrep], fields),
          spiraId: rowIdFieldInt(sheetData[rowToPrep], fields, fieldTypeEnums)
        },
          // create entry used to populate all relevant data for this row
          entry = {};
        // first check for errors
        var hasProblems = rowHasProblems(rowChecks, isUpdate);
        if (hasProblems) {
          entry.validationMessage = hasProblems;
          // if error free determine what field filtering is required - needed to choose type/subtype fields if subtype is present
        } else {
          var fieldsToFilter = relevantFields(rowChecks);
          entry = createEntryFromRow(sheetData[rowToPrep], model, fieldTypeEnums, artifactIsHierarchical, lastIndentPosition, fieldsToFilter, isUpdate, isComment, '');
          // FOR SUBTYPE ENTRIES add flag on entry if it is a subtype
          if (entry && fieldsToFilter === FIELD_MANAGEMENT_ENUMS.subType) {
            entry.isSubType = true;
          }
          // FOR HIERARCHICAL ARTIFACTS update the last indent position before going to the next entry to make sure relative indent is set correctly
          if (entry && artifactIsHierarchical) {
            lastIndentPosition = entry.indentPosition;
          }
        }
        //treating special Test Case update issues

        if (artifact.id == params.artifactEnums.testCases && entry.isSubType) {
          //if this is a testStep, check if the parent is valid
          var validParent = isValidParent(entriesForExport);
          if (!validParent) {
            //if the parent is not valid, mark that as an error
            entry = {};
            entry.validationMessage = 'Invalid TestCase parent. Please check your data.';
          }
        }

        //treating special Folders Creation

        if (artifact.id == params.artifactEnums.folders) {
          //We need to retrieve the Artifact we're creating the folder to and
          //use that to get the specific Parent Folder ID field
          var folderArtifact = params.artifacts.filter(function (artifact) {
            return artifact.id == entry.artifact;
          })[0];

          var newkey = "";

          //replacing Parent field, accordingly
          if (entry.Parent) {
            newkey = params.parentFolders[folderArtifact.id];
            Object.defineProperty(entry, newkey,
              Object.getOwnPropertyDescriptor(entry, "Parent"));

            //deleting the temporary (for internal reference only) Parent field
            delete entry.Parent;
          }

          //replacing folder ID field, accordingly

          /* newkey = params.IdFolders[folderArtifact.id];
           Object.defineProperty(entry, newkey, {
             //temp value
             value: 0
           })*/

          //deleting the temporary (for internal reference only) FolderId field
          //delete entry.FolderId;

        }
        if (entry) { entriesForExport.push(entry); }
      }
      else {
        //we are going to skip this entry
        var entry = {
          "skip": true
        }
        entriesForExport.push(entry);
      }
    }
  }
  return entriesForExport;
}


//2.2 Comment fields (if any) - Sent through the artifact comment API function (POST)
// loop to create artifact comment objects from each row taken from the spreadsheet
// vars needed: sheetData, artifact, fields, model, fieldTypeEnums, artifactIsHierarchical,
function createExportCommentEntries(sheetData, model, fieldTypeEnums, artifact, artifactIsHierarchical, artifactEntries, entriesLog) {
  var lastIndentPosition = null,
    entriesForExport = [],
    isComment = true;
  //in the function logics, comments are always update (we already have the ArtifactID)
  var isUpdate = true;

  for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
    // stop at the first row that is fully blank
    if (sheetData[rowToPrep].join("") === "") {
      break;
    } else {
      // create entry used to populate the comment field for this row
      var entry = {};
      // first check for errors
      var hasProblems = commentHasProblems(artifact);
      var sourceID = 0;
      if (artifactEntries[rowToPrep]) {
        if ('newId' in artifactEntries[rowToPrep]) {
          sourceID = artifactEntries[rowToPrep].newId;
        }
      }

      if (hasProblems) {
        entry.validationMessage = hasProblems;
        // if error free, create the entry
      } else {
        entry = createEntryFromRow(sheetData[rowToPrep], model, fieldTypeEnums, artifactIsHierarchical, lastIndentPosition, "", isUpdate, isComment, sourceID);
      }

      //checking if the parent has any problem
      //if so, skip this entry
      if (!entriesLog.entries[rowToPrep]) {
        entry.skip = true;
      }
      else if (entriesLog.entries[rowToPrep].error) {
        entry.skip = true;
      }

      entriesForExport.push(entry);
    }
  }
  return entriesForExport;
}

//2.3 Association fields (if any) - Sent through the artifact association API function (POST)
// loop to create artifact association objects from each row taken from the spreadsheet
// vars needed: sheetData, artifact, fields, model, fieldTypeEnums, artifactIsHierarchical
function createExportAssociationEntries(sheetData, model, fieldTypeEnums, artifact, artifactEntries, entriesLog) {
  var entriesForExport = [];

  //in the function logics, associations are always update (we already have the ArtifactIDs)
  for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
    // stop at the first row that is fully blank
    if (sheetData[rowToPrep].join("") === "") {
      break;
    } else {
      // create entry used to populate the comment field for this row
      var entry = [];
      var sourceID = 0;
      if (artifactEntries[rowToPrep]) {
        if ('newId' in artifactEntries[rowToPrep]) {
          sourceID = artifactEntries[rowToPrep].newId;
        }
      }

      // first check for errors
      var hasProblems = null; //associationHasProblems(sheetData[rowToPrep], model.fields);
      if (hasProblems) {
        entry.validationMessage = hasProblems;
        // if error free, create the entry
      } else {
        entry = createAssociationEntryFromRow(sheetData[rowToPrep], model, fieldTypeEnums, sourceID);
      }
      //checking if the parent has any problem
      //if so, skip this entry
      if (!entriesLog.entries[rowToPrep]) {
        entry.skip = true;
        entriesForExport.push(entry);
      }
      else if (entriesLog.entries[rowToPrep].error) {
        entry.skip = true;
        entriesForExport.push(entry);
      } else {
        //append each entry to the export object
        entry.forEach(function appendItems(item) {
          //double check to make sure this is a valid entry:
          if ((artifact.id == ART_ENUMS.requirements && item.DestArtifactId) || (item.skip)) {
            //if valid, append to the array
            entriesForExport.push(item);
          }
          else if ((artifact.id == ART_ENUMS.testCases && item.RequirementId) || (item.skip)) {
            //if valid, append to the array
            entriesForExport.push(item);
          }
          else if ((artifact.id == ART_ENUMS.testCases && item.ReleaseId) || (item.skip)) {
            //if valid, append to the array
            entriesForExport.push(item);
          }
          else if ((artifact.id == ART_ENUMS.testCases && item.TestSetId) || (item.skip)) {
            //if valid, append to the array
            entriesForExport.push(item);
          }
        });
      }
    }
  }
  return entriesForExport;
}


// 3. FOR GOOGLE ONLY: GET READY TO SEND DATA TO SPIRA + 4. ACTUALLY SEND THE DATA
// check we have some entries and with no errors
// Create and show a message to tell the user what is going on
function sendExportEntriesGoogle(entriesForExport, commentEntriesForExport, associationEntriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, validations) {
  if (!entriesForExport.length && !commentEntriesForExport.length && !associationEntriesForExport.length) {
    popupShow('There are no entries to send to Spira', 'Check Sheet')
    return "nothing to send";
  } else {
    popupShow('Preparing to send...', 'Progress');

    // create required variables for managing responses for sending data to spiraplan
    var log = {
      errorCount: 0,
      successCount: 0,
      doNotContinue: false,
      // set var for parent - used to designate eg a test case so it can be sent with the test step post
      parentId: -1,
      entriesLength: entriesForExport.length,
      entries: []
    };

    // loop through objects to send and update the log
    function sendSingleEntry(i) {
      // set the correct parentId for hierarchical artifacts
      // set before launching the API call as we need to look back through previous entries
      if (artifact.hierarchical) {
        log.parentId = getHierarchicalParentId(entriesForExport[i].indentPosition, log.entries);
      }
      //if we not have a parent ID yet, set the correct parentId artifact for subtypes (needed for POST URL)
      if (isUpdate && Number(log.parentId) == -1 && artifact.hasSubType && entriesForExport[i].isSubType) {
        log.parentId = getAssociationParentInfo(entriesForExport, i, artifact.id, "id");
        log.parentName = getAssociationParentInfo(entriesForExport, i, artifact.id, "name");
        log.parentActive = getAssociationParentInfo(entriesForExport, i, artifact.id, "active");
        //let the child object to hold the parent ID field
        entriesForExport[i][ART_PARENT_IDS[artifact.id]] = log.parentId;

        if ('parentName' in log) {
          entriesForExport[i].parentName = log.parentName;
        }
        if ('parentActive' in log) {
          entriesForExport[i].parentActive = log.parentActive;
        }
      }
      var logResponse = manageSendingToSpira(entriesForExport[i], model.user, model.currentProject.id, model.currentTemplate.id, artifact, fields, fieldTypeEnums, log.parentId, isUpdate, false, false);
      // update the parent ID for a subtypes based on the successful API call
      //if (artifact.hasSubType && !entriesForExport[i].isSubType) {
      if (artifact.hasSubType) {
        //BRUNO
        log.parentId = logResponse.newId;
      }
      log = processSendToSpiraResponse(i, logResponse, entriesForExport, artifact, log, false, false);

      if (logResponse.error && artifact.hierarchical && !isUpdate) {
        // break out of the recursive loop
        log.doNotContinue = true;
      }
      //reset the variable to the next position
      if (isUpdate) {
        log.parentId = -1;
      }
    }

    // loop through comment objects to send
    function sendSingleCommentEntry(j) {
      if (!validations.includes(commentEntriesForExport[j].ArtifactId) && !(commentEntriesForExport[j].skip)) { //if this is a valid entry
        var logResponse = manageSendingToSpira(commentEntriesForExport[j], model.user, model.currentProject.id, model.currentTemplate.id, artifact, fields, fieldTypeEnums, log.parentId, isUpdate, true, false);
        // update the parent ID for a subtypes based on the successful API call
        if (artifact.hasSubType) {
          log.parentId = logResponse.parentId;
        }
        log = processSendToSpiraResponse(j, logResponse, commentEntriesForExport, artifact, log, true, false);
      }
    }

    // loop through association objects to send
    function sendSingleAssociationEntry(k) {
      if (!validations.includes(associationEntriesForExport[k].SourceArtifactId) && !(associationEntriesForExport[k].skip)) { //if this is a valid entry
        var logResponse = manageSendingToSpira(associationEntriesForExport[k], model.user, model.currentProject.id, model.currentTemplate.id, artifact, fields, fieldTypeEnums, log.parentId, isUpdate, false, true)
        // update the parent ID for a subtypes based on the successful API call
        if (artifact.hasSubType) {
          log.parentId = logResponse.parentId;
        }

        log = processSendToSpiraResponse(k, logResponse, associationEntriesForExport, artifact, log, false, true);
      }
    }


    // 4. SEND DATA TO SPIRA AND MANAGE RESPONSES
    // KICK OFF THE FOR LOOP (IE THE FUNCTION ABOVE) HERE
    // We use a function rather than a loop so that we can more readily use promises to chain things together and make the calls happen synchronously
    // we need the calls to be synchronous because we need to do the status and ID of the preceding entry for hierarchical artifacts
    for (var i = 0; i < entriesForExport.length; i++) {
      if (!entriesForExport[i].skip) {
        if (!log.doNotContinue) {
          log = checkSingleEntryForErrors(entriesForExport[i], log, artifact, isUpdate);
          if (log.entries.length && log.entries[i] && log.entries[i].error) {
            // do nothing
          } else {
            sendSingleEntry(i);
          }
        }
      }
      else {
        var skip = {
          "skip": true
        };
        log.entries.push(skip);
        log.successCount++;

      }
    }

    //then, send comments for the artifact (if any)
    for (var j = 0; j < commentEntriesForExport.length; j++) {
      //make sure we don't have any error and we have the comment populated
      if (Object.keys(commentEntriesForExport[j]).length > 1) {
        sendSingleCommentEntry(j);
      }
    }

    //if any association is available, send them
    for (var k = 0; k < associationEntriesForExport.length; k++) {
      //make sure we don't have any error and we have the comment populated
      if (Object.keys(associationEntriesForExport[k]).length > 1) {
        sendSingleAssociationEntry(k);
      }
    }
    // review all activity and set final status
    log.status = setFinalStatus(log);
    return log;
  }
}


// 3. FOR EXCEL ONLY: GET READY TO SEND DATA TO SPIRA + 4. ACTUALLY SEND THE DATA
// DIFFERENT TO GOOGLE: this uses js ES6 a-sync and a-wait for its function and subfunction
// check we have some entries and with no errors
// Create and show a message to tell the user what is going on
async function sendExportEntriesExcel(entriesForExport, commentEntriesForExport, associationEntriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context, isUpdate, validations) {
  if (!entriesForExport.length && !commentEntriesForExport.length && !associationEntriesForExport.length) {
    popupShow('There are no entries to send to Spira', 'Check Sheet')
    return "nothing to send";
  } else {
    popupShow('Starting to update...', 'Progress');
    // create required variables for managing responses for sending data to spiraplan
    var log = {
      errorCount: 0,
      successCount: 0,
      doNotContinue: false,
      // set var for parent - used to designate eg a test case so it can be sent with the test step post
      parentId: -1,
      entriesLength: entriesForExport.length,
      entries: []
    };


    // loop through objects to send and update the log
    async function sendSingleEntry(i) {
      // set the correct parentId for hierarchical artifacts
      // set before launching the API call as we need to look back through previous entries
      if (artifact.hierarchical) {
        log.parentId = getHierarchicalParentId(entriesForExport[i].indentPosition, log.entries);
      }
      //if we not have a parent info yet, set the correct parentId, parentName and parentActive fields for subtypes (needed for POST URL)
      if (isUpdate && Number(log.parentId) == -1 && artifact.hasSubType && entriesForExport[i].isSubType) {
        log.parentId = getAssociationParentInfo(entriesForExport, i, artifact.id, "id");
        log.parentName = getAssociationParentInfo(entriesForExport, i, artifact.id, "name");
        log.parentActive = getAssociationParentInfo(entriesForExport, i, artifact.id, "active");
        //let the child object to hold the parent ID field
        entriesForExport[i][ART_PARENT_IDS[artifact.id]] = log.parentId;

        if ('parentName' in log) {
          entriesForExport[i].parentName = log.parentName;
        }
        if ('parentActive' in log) {
          entriesForExport[i].parentActive = log.parentActive;
        }
      }
      await manageSendingToSpira(entriesForExport[i], model.user, model.currentProject.id, model.currentTemplate.id, artifact, fields, fieldTypeEnums, log.parentId, isUpdate, false, false)
        .then(function (response) {
          // update the parent ID for a subtypes based on the successful API call
          if (artifact.hasSubType) {
            log.parentId = response.parentId;
          }
          log = processSendToSpiraResponse(i, response, entriesForExport, artifact, log, false, false);
          if (response.error && artifact.hierarchical && !isUpdate) {
            // break out of the recursive loop
            log.doNotContinue = true;
          }
        })
      //reset the variable to the next position
      if (isUpdate) {
        log.parentId = -1;
      }
    }

    // loop through comment objects to send
    async function sendSingleCommentEntry(j) {
      if (!validations.includes(commentEntriesForExport[j].ArtifactId)) { //if this is a valid entry
        await manageSendingToSpira(commentEntriesForExport[j], model.user, model.currentProject.id, model.currentTemplate.id, artifact, fields, fieldTypeEnums, log.parentId, isUpdate, true, false)
          .then(function (response) {
            // update the parent ID for a subtypes based on the successful API call
            if (artifact.hasSubType) {
              log.parentId = response.parentId;
            }
            log = processSendToSpiraResponse(j, response, commentEntriesForExport, artifact, log, true, false);
          })
      }
    }

    // loop through association objects to send
    async function sendSingleAssociationEntry(k) {
      if (!validations.includes(associationEntriesForExport[k].SourceArtifactId)) { //if this is a valid entry
        await manageSendingToSpira(associationEntriesForExport[k], model.user, model.currentProject.id, model.currentTemplate.id, artifact, fields, fieldTypeEnums, log.parentId, isUpdate, false, true)
          .then(function (response) {
            // update the parent ID for a subtypes based on the successful API call
            if (artifact.hasSubType) {
              log.parentId = response.parentId;
            }
            log = processSendToSpiraResponse(k, response, associationEntriesForExport, artifact, log, false, true);
          })
      }
    }


    // 4. SEND DATA TO SPIRA AND MANAGE RESPONSES
    // KICK OFF THE FOR LOOP (IE THE FUNCTION ABOVE) HERE
    // We use a function rather than a loop so that we can more readily use promises to chain things together and make the calls happen synchronously
    // we need the calls to be synchronous because we need to do the status and ID of the preceding entry for hierarchical artifacts
    //first, send standard and custom artifact properties
    for (var i = 0; i < entriesForExport.length; i++) {
      if (!entriesForExport[i].skip) {
        if (!log.doNotContinue) {
          log = checkSingleEntryForErrors(entriesForExport[i], log, artifact);
          if (log.entries.length && log.entries[i] && log.entries[i].error) {
            // do nothing 
          } else {
            await sendSingleEntry(i);
          }
        }
      }
      else {
        var skip = {
          "skip": true
        };
        log.entries.push(skip);
        log.successCount++;

      }
    }

    //then, send comments for the artifact (if any)
    for (var j = 0; j < commentEntriesForExport.length; j++) {
      //make sure we don't have any error and we have the comment populated
      if (Object.keys(commentEntriesForExport[j]).length > 1) {
        await sendSingleCommentEntry(j);
      }
    }

    //if any association is available, send them
    for (var k = 0; k < associationEntriesForExport.length; k++) {
      //make sure we don't have any error and we have the comment populated
      if (Object.keys(associationEntriesForExport[k]).length > 1) {
        await sendSingleAssociationEntry(k);
      }
    }

    // review all activity and set final status
    log.status = setFinalStatus(log);
    // call the final function here - so we know that it is only called after the recursive function above (ie all posting) has ended
    return log;
  }
}


// 5. SET MESSAGES AND FORMATTING ON SHEET
function updateSheetWithExportResults(entriesLog, commentsLog, associationsLog, entriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, isUpdate, context) {
  var bgColors = [],
    notes = [],
    values = [];
  var associationCounter = 0;
  // first handle cell formatting
  for (var row = 0; row < entriesForExport.length; row++) {
    var rowBgColors = [],
      rowNotes = [],
      rowValues = [],
      associationNote = '';
    for (var col = 0; col < fields.length; col++) {
      var bgColor,
        note = null,
        value = sheetData[row][col];

      // we may have more rows than entries - because the entries can be stopped early (eg when an error is found on a hierarchical artifact)
      if (entriesLog != null) {
        if (entriesLog.entries.length > row) {
          //check for comments error
          if (fields[col].isComment && commentsLog != null) {
            if (commentsLog.entries[row]) {
              if (commentsLog.entries[row].error) {
                bgColor = model.colors.warning;
              }
            }
          }
          //check for associations error
          else if (fields[col].association && associationsLog != null && associationsLog.entries != null) {

            //we've reached an association field. First, we need to counter how many associations we have here

            var associationText = sheetData[row][col];
            var associationSize = 0;

            //get the association field size
            if ((associationText + '') != '') {
              associationText = (associationText + '').replace('.', ',');
              associationText = associationText.replace(/\s/g, '');
              if (!isInt(associationText)) {
                { //in this case, we must have a comma-separated string
                  var associationIds = (associationText + '').split(',');
                  var associationSize = associationIds.length;
                }
              }
              else {
                associationSize = 1;
              }
            }

            for (var index = 0; index < associationSize; index++) {
              if (associationsLog.entries[associationCounter]) {
                if (associationsLog.entries[associationCounter].error) {
                  bgColor = model.colors.warning;
                  associationNote += params.associationTextLabels[fields[col].association] + ' ' + '(' + (index + 1) + ')' + ' Error: ' + associationsLog.entries[associationCounter].message + ' ';
                }
              }
              //keeping track of associations - they're not directly related to the row number
              associationCounter++;
            }
          }
          else {
            // first handle when we are dealing with data that has been sent to Spira
            var isSubType = (entriesLog.entries[row].details && entriesLog.entries[row].details.entry && entriesLog.entries[row].details.entry.isSubType) ? entriesLog.entries[row].details.entry.isSubType : false;
            bgColor = setFeedbackBgColor(sheetData[row][col], entriesLog.entries[row].error, fields[col], fieldTypeEnums, artifact, model.colors);
            value = setFeedbackValue(sheetData[row][col], entriesLog.entries[row].error, fields[col], fieldTypeEnums, entriesLog.entries[row].newId || "", isSubType, col);
            note = setFeedbackNote(sheetData[row][col], entriesLog.entries[row].error, fields[col], fieldTypeEnums, entriesLog.entries[row].message, value, isUpdate);
          }
        }
      }
      if (IS_GOOGLE) {
        rowBgColors.push(bgColor);
        if (associationNote.length > 0) {
          rowNotes.push(associationNote);
        }
        else {
          rowNotes.push(note);
        }
        rowValues.push(value);
      } else {
        var cellRange = sheet.getCell(row + 1, col);
        if (note) rowNotes.push(note);
        if (associationNote.length > 0) {
          var subRange = sheet.getCell(row + 1, params.resultColumns[model.currentArtifact.name]);
          subRange.values = [[associationNote]];
        }
        if (bgColor) {
          cellRange.set({ format: { fill: { color: bgColor } } });
        }
        if (value) {
          cellRange.values = [[value]];
          context.sync();
        }
      }
    }
    if (IS_GOOGLE) {
      bgColors.push(rowBgColors);
      notes.push(rowNotes);
      values.push(rowValues);

      // for Excel we can't pass in arrays of data for values, but we still take action here for notes - because Excel API does not allow the addition of comments
    } else {
      var rowFirstCell = sheet.getCell(row + 1, 0);
      if (rowNotes.length > 0) {
        rowFirstCell.set({ format: { fill: { color: model.colors.warning } } });
        rowFirstCell.values = [[rowNotes.join()]];
      }
    }
  }
  if (IS_GOOGLE) {
    sheetRange.setBackgrounds(bgColors).setNotes(notes).setValues(values);
    return entriesLog;
  } else {
    if (context != null) {
      return context.sync().then(function () { return entriesLog; });
    }
    else {
      return entriesLog;
    }
  }
}

function checkSingleEntryForErrors(singleEntry, log, artifact, isUpdate) {
  var response = {};
  // skip if there was an error validating the sheet row
  if (singleEntry.validationMessage) {
    response.error = true;
    response.message = singleEntry.validationMessage;
    log.errorCount++;

    // stop if the artifact is hierarchical because we don't know what side effects there could be to any further items.
    if (artifact.hierarchical && !isUpdate) {
      log.doNotContinue = true;
      response.message += " - no further entries were sent.";
      // make sure to push the response so that the client can process error message
      log.entries.push(response);
      // we do not call this function again with i++ so that we effectively break out of the loop
      // review all activity and set final status
      log.status = log.errorCount ? (log.errorCount == log.entriesLength ? STATUS_ENUM.allError : STATUS_ENUM.someError) : STATUS_ENUM.allSuccess;
    } else {
      log.entries.push(response);
    }
    // skip if a sub type row does not have a parent to hook to
  } else if (singleEntry.isSubType && !log.parentId) {
    response.error = true;
    response.message = "can't add a child type when there is no corresponding parent type";
    log.errorCount++;
    log.entries.push(response);
  }
  return log;
}


// utility function to set final status of the log
// @param: log - log object
// returns enum for the final status
function setFinalStatus(log) {
  if (log.errorCount) {
    //check if any error message is about data validation etc - if not then all the message are about the entry already existing in Spira (where the message is the INT of the id)
    let logEntriesOnlyAboutIds = true;
    const logMessages = log.entries.filter(x => x.message);
    if (logMessages.length) {
      for (let index = 0; index < logMessages.length; index++) {
        if (!Number.isInteger(parseInt(logMessages[index].message))) {
          logEntriesOnlyAboutIds = false;
          break;
        }
      }
    }

    if (logEntriesOnlyAboutIds) {
      return STATUS_ENUM.existingEntries;
    } else if (log.errorCount == log.entriesLength) {
      return STATUS_ENUM.allError;
    } else {
      return STATUS_ENUM.someError;
    }
  } else {
    return STATUS_ENUM.allSuccess;
  }
}



// function that reviews a specific cell against it's field and errors for providing UI feedback on errors
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldTypeEnums - enum information about field types
// @param: artifact - the currently selected artifact
// @param: colors - object of colors to use based on different conditions
function setFeedbackBgColor(cell, error, field, fieldTypeEnums, artifact, colors) {
  if (error) {
    // if we have a validation error, we can highlight the relevant cells if the art has no sub type
    if (!artifact.hasSubType) {
      if (field.required && cell === "") {
        return colors.warning;
      } else {
        // keep original formatting
        if (field.type == fieldTypeEnums.subId || field.type == fieldTypeEnums.id || field.unsupported) {
          return colors.warning;
        } else {
          return null;
        }
      }

      // otherwise highlight the whole row as we don't know the cause of the problem
    } else {
      return colors.warning;
    }

    // no errors
  } else {
    // keep original formatting
    if (field.type == fieldTypeEnums.subId || field.type == fieldTypeEnums.id || field.unsupported) {
      return colors.bgReadOnly;
    } else {
      return null;
    }
  }
}



// function that reviews a specific cell against it's field and sets any notes required
// currently only adds error message as note to ID field
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldTypeEnums - enum information about field types
// @param: message - relevant error message from the entry for this row
// @param: value - original ID of the artifact to be kept in case of an error
function setFeedbackNote(cell, error, field, fieldTypeEnums, message, value, isUpdate) {
  // handle entries with errors - add error notes into ID field
  if (error && field.type == fieldTypeEnums.id) {
    //invalid new rows always have -1 in the ID field
    if (value == "-1" && isUpdate) {
      //however, special subtype IDs can be blank, that's not an error
      if (!SUBTYPE_IDS.includes(field.field)) {
        return "Error: you can't send new data to Spira when updating. Please use the 'Send data to Spira' option."
      }
      else {
        return ',Error: ' + message;
      }
    }
    else if (!isUpdate && cell != '') {
      return ' Error: row previously added to Spira. Skipped to avoid duplicates.';
    }
    else {
      return value + ',' + message;
    }
  } else {
    return null;
  }
}

// function that updates id fields with new values, otherwise returns existing value
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldTypeEnums - enum information about field types
// @param: newId - int that is the newly created Id for this row
// @param: isSubType - bool if row is subtype or not - false on error as there will be no id to add anyway
function setFeedbackValue(cell, error, field, fieldTypeEnums, newId, isSubType, col) {
  // when there is an error we don't change any of the cell data
  if (error && !(col == 0 && cell == -1)) {
    return cell;

    // handle successful entries - ie add ids into right place
  } else {
    var newIdToEnter = newId || "";
    if (!isSubType && field.type == fieldTypeEnums.id) {
      return newIdToEnter;
    } else if (isSubType && field.type == fieldTypeEnums.subId) {
      return newIdToEnter;
    } else if (isSubType && field.type == fieldTypeEnums.id && col == 0 && cell == -1) {
      //this fix the visual bug for test steps
      return '';
    } else {
      return cell;
    }
  }
}



// on determining that an entry should be sent to Spira, this function handles calling the API function, and parses the data on both success and failure
// @param: entry - object of the specific entry in format ready to attach to body of API request
// @param: parentId - int of the parent id for this specific loop - used for attaching subtype children to the right parent artifact
// @param: artifact - object of the artifact being used here to help manage what specific API call to use
// @param: user - user object for API call authentication
// @param: projectId - int of project id for API call
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldTypeEnums - object of all field types with enums
// @param: isUpdate - flag to indicate if this is an update operation
// @param: isComment - flag to indicate if this is a comment creation operation
function manageSendingToSpira(entry, user, projectId, templateId, artifact, fields, fieldTypeEnums, parentId, isUpdate, isComment, isAssociation) {
  var data,
    // make sure correct artifact ID is sent to handler (ie type vs subtype)
    artifactTypeIdToSend = entry.isSubType ? artifact.subTypeId : artifact.id,
    // set output parent id here so we know this function will always return a value for this
    output = {
      parentId: parentId,
      entry: entry,
      artifact: {
        artifactId: artifactTypeIdToSend,
        artifactObject: artifact
      }
    };

  // send object to relevant artifact post service
  if (IS_GOOGLE) {
    if (!isComment && !isAssociation) {
      if (!isUpdate || (isUpdate && artifact.allowsCreateOnUpdate && entry.createOnUpdate)) {
        data = postArtifactToSpira(entry, user, projectId, templateId, artifactTypeIdToSend, parentId);
      }
      else {
        data = putArtifactToSpira(entry, user, projectId, templateId, artifactTypeIdToSend, parentId);
      }
      // save data for logging to client
      output.entry = entry;
      output.httpCode = (data && data.getResponseCode()) ? data.getResponseCode() : "notSent";
      // parse the data if we have a success
      if (output.httpCode == 200) {
        output.fromSpira = JSON.parse(data.getContentText());
      }
      else {
        output.error = true;
        if (data && data.getContentText()) {
          output.errorMessage = data.getContentText();
        } else {
          output.errorMessage = "send attempt failed";
        }
        // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
        if (artifact.hasSubType && !entry.isSubType) {
          output.parentId = 0;
        }
      }

      if (!output.error) {
        // get the id/subType id of the updated artifact
        var artifactIdField = getIdFieldName(fields, fieldTypeEnums, entry.isSubType);

        //special case for folders
        if (artifactTypeIdToSend == ART_ENUMS.folders) {
          //
          var idField = params.IdFolders[entry.artifact];
          output.newId = output.fromSpira[idField];
        }
        else if (output.fromSpira && output.fromSpira != null) {
          //get the just-created ID from the server response
          output.newId = output.fromSpira[artifactIdField];
        }
        else {
          //just repeat the id - it's the same 
          output.newId = entry[artifactIdField];
        }

        // repeats the output parent ID only if the artifact has a subtype and this entry is NOT a subtype
        if (artifact.hasSubType && !entry.isSubType) {
          output.parentId = parentId;
        }
      }
      return output;
    }

    else if (isComment) {
      try {
        //comment creation
        var serverResponse = postCommentToSpira(entry, user, projectId, artifactTypeIdToSend);

        output.fromSpira = serverResponse.body;
        return output;
      }
      catch (error) {
        //we have an error - so set the flag and the message
        output.error = true;
        if (error) {
          output.errorMessage = error;
        } else {
          output.errorMessage = "send attempt failed";
        }
        // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
        if (artifact.hasSubType && !entry.isSubType) {
          output.parentId = 0;
        }
        return output;
      }
    }
    else {
      try {
        //isAssociation
        var serverResponse = postAssociationToSpira(entry, user, projectId, artifactTypeIdToSend, null);
        // save data for logging to client
        output.entry = entry;
        output.httpCode = (serverResponse && serverResponse.getResponseCode()) ? serverResponse.getResponseCode() : "notSent";
        // parse the data if we have a success
        if (output.httpCode == 200) {
          output.fromSpira = JSON.parse(serverResponse.getContentText());
        }
        else {
          output.error = true;
          if (serverResponse && serverResponse.getContentText()) {
            output.errorMessage = serverResponse.getContentText();
          } else {
            output.errorMessage = "send attempt failed";
          }
        }
        return output;
      }
      catch {
        //we have an error - so set the flag and the message
        output.error = true;
        output.errorMessage = "unknown error";

        // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
        if (artifact.hasSubType && !entry.isSubType) {
          output.parentId = 0;
        }
        return output;
      }
    }
  } else {
    //Excel
    if (!isComment && !isAssociation) {
      //checks for operation - also checks for special case, when creating while updating is allowed
      if (!isUpdate || (isUpdate && artifact.allowsCreateOnUpdate && entry.createOnUpdate)) {
        return postArtifactToSpira(entry, user, projectId, templateId, artifactTypeIdToSend, parentId)
          .then(function (response) {
            output.fromSpira = response.body;
            // get the id/subType id of the newly created artifact
            var artifactIdField = getIdFieldName(fields, fieldTypeEnums, entry.isSubType);
            //special case for folders
            if (artifactTypeIdToSend == ART_ENUMS.folders) {
              //
              var idField = params.IdFolders[entry.artifact];
              output.newId = output.fromSpira[idField];
            }
            else {
              output.newId = output.fromSpira[artifactIdField];
            }
            // update the output parent ID to the new id only if the artifact has a subtype and this entry is NOT a subtype
            if (artifact.hasSubType && !entry.isSubType) {
              output.parentId = output.newId;
            }
            return output;
          })
          .catch(function (error) {
            //we have an error - so set the flag and the message
            output.error = true;
            if (error) {
              output.errorMessage = error;
            } else {
              output.errorMessage = "send attempt failed";
            }

            // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
            if (artifact.hasSubType && !entry.isSubType) {
              output.parentId = 0;
            }
            return output;
          });

      }
      else {
        return putArtifactToSpira(entry, user, projectId, templateId, artifactTypeIdToSend, parentId)
          .then(function (response) {
            var errorStatus = response.error;
            if (!errorStatus) {
              // get the id/subType id of the updated artifact
              var artifactIdField = getIdFieldName(fields, fieldTypeEnums, entry.isSubType);
              //just repeat the id - it's the same 
              output.newId = entry[artifactIdField];
              // repeats the output parent ID only if the artifact has a subtype and this entry is NOT a subtype
              if (artifact.hasSubType && !entry.isSubType) {
                output.parentId = parentId;
              }
              return output;

            }
          })
          .catch(function (error) {
            //we have an error - so set the flag and the message
            output.error = true;
            if (error) {
              output.errorMessage = error;
            } else {
              output.errorMessage = "update attempt failed";
            }

            // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
            if (artifact.hasSubType && !entry.isSubType) {
              output.parentId = 0;
            }
            return output;
          });
      }
    }
    else if (isComment) {
      //comment creation
      return postCommentToSpira(entry, user, projectId, artifactTypeIdToSend)
        .then(function (response) {
          output.fromSpira = response.body;
          return output;
        })
        .catch(function (error) {
          //we have an error - so set the flag and the message
          output.error = true;
          if (error) {
            output.errorMessage = error;
          } else {
            output.errorMessage = "send attempt failed";
          }

          // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
          if (artifact.hasSubType && !entry.isSubType) {
            output.parentId = 0;
          }
          return output;
        });
    }
    else {
      //isAssociation
      return postAssociationToSpira(entry, user, projectId, artifactTypeIdToSend)
        .then(function (response) {
          output.fromSpira = response.body;
          return output;
        })
        .catch(function (error) {
          //we have an error - so set the flag and the message
          output.error = true;
          if (error) {
            output.errorMessage = error;
          } else {
            output.errorMessage = "send attempt failed";
          }

          // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
          if (artifact.hasSubType && !entry.isSubType) {
            output.parentId = 0;
          }
          return output;
        });

    }

  }
}

// returns the correct parentId for the relevant indent position by looping back through the list of entries
// returns -1 if no match found
// @param: indent - int of the indent position to retrieve the parent for
// @param: previousEntries - object containing all successfully sent entries - with, if a hierarchical artifact, a hierarchy info object
function getHierarchicalParentId(indent, previousEntries) {
  // if there is no indent/ set to initial indent we return out immediately 
  if (indent === 0 || !previousEntries.length) {
    return -1;
  }
  for (var i = previousEntries.length - 1; i >= 0; i--) {
    // when the indent is greater - means we are indenting, so take the last array item and return out
    // check for presence of correct objects in item - should exist, as otherwise error should be thrown
    if (previousEntries[i].details && previousEntries[i].details.hierarchyInfo && previousEntries[i].details.hierarchyInfo.indent < indent) {
      return previousEntries[i].details.hierarchyInfo.id;
    }
  }
  return -1;
}

// returns the correct item for the relevant group position by looping back through the list of entries
// returns an invalid value if no match found
// @param: entriesForExport - the whole list of artifacts in the spreadsheet
// @param: i - target artifact position to look for
// @param: artifactId - artifact type
// @mode: what to look for 
function getAssociationParentInfo(entriesForExport, i, artifactId, mode) {

  if (mode == "id") {
    // we need to know the artifact type to look for specific parent ID fields
    if (artifactId == ART_ENUMS.testCases || artifactId == ART_ENUMS.customLists) {
      //look for parents in the previous rows
      for (i; i > 0; i--) {
        //if we found the specific parentID for the artifact type, return it
        if (entriesForExport[i - 1].hasOwnProperty(ART_PARENT_IDS[artifactId])) {
          return entriesForExport[i - 1][ART_PARENT_IDS[artifactId]];
        }
      }
      return -1;
    }
  }
  else if (mode == "name") {
    // we need to know the artifact type to look for specific parent Name fields
    if (artifactId == ART_ENUMS.customLists) {
      //look for parents in the previous rows
      for (i; i > 0; i--) {
        //if we found the specific parentID for the artifact type, return it
        if (entriesForExport[i - 1].hasOwnProperty("Name") && !entriesForExport[i - 1].isSubType) {
          return entriesForExport[i - 1].Name;
        }
      }
      return "";
    }
  }
  else if (mode == "active") {
    // we need to know the artifact type to look for specific parent Name fields
    if (artifactId == ART_ENUMS.customLists) {
      //look for parents in the previous rows
      for (i; i > 0; i--) {
        //if we found the specific parentID for the artifact type, return it
        if (entriesForExport[i - 1].hasOwnProperty("Active") && !entriesForExport[i - 1].isSubType) {
          return entriesForExport[i - 1].Active;
        }
      }
      return false;
    }
  }
  return null;
}


// returns an int of the total number of required fields for the passed in artifact
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
// @param: forSubType - bool to determine whether to check for sub type required fields (true), or not - defaults to false
function countRequiredFieldsByType(fields, forSubType) {
  var count = 0;
  for (var i = 0; i < fields.length; i++) {
    if (forSubType != "undefined" && forSubType) {
      if (fields[i].requiredForSubType) {
        count++;
      }
    } else if (fields[i].required) {
      count++;
    }
  }
  return count;
}



// check to see if a row of data has entries for all required fields
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
function rowCountRequiredFieldsByType(row, fields, forSubType) {
  var count = 0;
  for (var i = 0; i < row.length; i++) {
    if (forSubType != "undefined" && forSubType) {
      if (fields[i].requiredForSubType && row[i]) {
        count++;
      }
    } else if ((fields[i].required && row[i]) || (fields[i].required && row[i] == '0')) {
      count++;
    }

  }
  return count;
}



// check to see if a row for an artifact with a subtype has a field that can't be present if subtype fields are filled in
// this can be useful to make sure that one field - eg Test Case Name would make sure a test step is not created to avoid any confusion
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
function rowBlocksSubType(row, fields) {
  var result = false;
  for (var column = 0; column < row.length; column++) {
    if (fields[column].forbidOnSubType && row[column]) {
      result = true;
    }
  }
  return result;
}



// check to see if a row for an artifact has any id field filled in with an int (not a string - a string could mean a different error message was previously added with Excel)
// returns false if id field is not an int, returns the ID in the cell if one is present as an INT (ie send back the Spira ID)
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldTypeEnums - object of all field types with enums
function rowIdFieldInt(row, fields, fieldTypeEnums) {
  var result = false;
  for (var column = 0; column < row.length; column++) {
    const cellIsIdField = fields[column].type === fieldTypeEnums.id || fields[column].type === fieldTypeEnums.subId;
    if (cellIsIdField && Number.isInteger(parseInt(row[column]))) {
      //set the result to the value of the ID so that the row is skipped but in the UI it looks the same - ie has the correct ID etc
      result = row[column];
      break;
    }
  }
  return result;
}



// checks to see if the row is valid - ie required fields present and correct as expected
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: rowChecks - object with different properties for different checks required
function rowHasProblems(rowChecks, isUpdate) {
  var problem = null;
  // if the row already exists in Spira then do not carry out any other problem analysis
  if (rowChecks.spiraId && !isUpdate) {
    problem = rowChecks.spiraId;
    // if a new entry, carry out problem analysis
  } else if (!rowChecks.hasSubType && rowChecks.countRequiredFields < rowChecks.totalFieldsRequired) {
    problem = "Fill in all required fields";
  } else if (rowChecks.hasSubType) {
    if (rowChecks.countSubTypeRequiredFields < rowChecks.totalSubTypeFieldsRequired && !rowChecks.countRequiredFields) {
      problem = "Fill in all required fields";
    } else if (rowChecks.countRequiredFields < rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
      problem = "Fill in all required fields";
    } else if (rowChecks.countSubTypeRequiredFields == rowChecks.totalSubTypeFieldsRequired && (rowChecks.countRequiredFields == rowChecks.totalFieldsRequired || rowChecks.subTypeIsBlocked)) {
      problem = "It is unclear what artifact this is intended to be";
    }
  }
  return problem;
}

// checks to see if a coment has any problem related to it
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: f - object with different properties for different checks required
function commentHasProblems(artifact) {
  var problem = null;
  // SubTypes can't have comments
  if (artifact.isSubType) {
    problem = "Comment not allowed for this artifact type.";
  }
  return problem;
}

//checks if the text corresponds to a valid string
function isInt(value) {
  return !isNaN(value) && (function (x) { return (x | 0) === x; })(parseFloat(value))
}

//checks if the text we have for associations is a valid number OR a valid comma-separated string
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: - row to inspect - object with different properties for different checks required
//          - artifact - the artifact object model
function associationHasProblems(row, artifact) {

  var problem = null;

  var associationText = row.filter(function (item, index) {
    if (artifact[index].association) {
      return row[index];
    }
  })

  if ((associationText + '') != '') {
    //depending on the user's language, the console can misinterpret commas as points
    associationText = (associationText + '').replace('.', ',');
    if (!isInt(associationText)) {
      { //in this case, we must have a comma-separated string
        var associationIds = (associationText + '').split(',');
        var associationCount = associationIds.length;
        if (associationCount == 1) {
          problem = "Artifact Association data wrong format."
        }
        else {
          //check if every chunk is an integer
          associationIds.forEach(function (item) {
            if (!isInt(item)) problem = "Artifact Association data wrong format.";
          });

        }
      }
    }

  }
  return problem;
}

// based on field type and conditions, determines what fields are required for a given row
// e.g. all fields is default and standard, if a subtype is present (eg test step) - should it send only the main type or the sub type fields
// returns a int representing the relevant enum value
// @ param: rowChecks - object with different properties for different checks required
function relevantFields(rowChecks) {
  var fields = FIELD_MANAGEMENT_ENUMS.all;
  if (rowChecks.hasSubType) {
    if (rowChecks.countRequiredFieldsFilled == rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
      fields = FIELD_MANAGEMENT_ENUMS.standard;
    } else if (rowChecks.countSubTypeRequiredFields == rowChecks.totalSubTypeFieldsRequired && !(rowChecks.countRequiredFields == rowChecks.totalFieldsRequired || rowChecks.subTypeIsBlocked)) {
      fields = FIELD_MANAGEMENT_ENUMS.subType;
    }
  }

  return fields;
}



// function creates a correctly formatted artifact object ready to send to Spira
// it works through each field type to validate and parse the values so object is in correct form
// any field that does not pass validation receives a null value
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: model - full model with info about fields, dropdowns, users, etc
// @param: fieldTypeEnums - object of all field types with enums
// @param: artifactIsHierarchical - bool to tell function if this artifact has hierarchy (eg RQ and RL)
// @param: lastIndentPosition - int used for calculating relative indents for hierarchical artifacts
// @param: fieldsToFilter - enum used for selecting fields to not add to object - defaults to using all if omitted
// @param: isUpdate - bool to flag id this is an update operation. If false, it is a creation operation
// @param: isComment - bool to flag if we will return a comment entry (true) or custom/standard (false)
function createEntryFromRow(row, model, fieldTypeEnums, artifactIsHierarchical, lastIndentPosition, fieldsToFilter, isUpdate, isComment, sourceId) {

  var fields = model.fields;
  var entry = {};
  var missingSubId = false;
  //populate 'entry' object accordingly - include custom properties array here to avoid it being undefined later if needed
  if (!isComment) {
    var entry = {
      "CustomProperties": []
    }
  }

  //we need to turn an array of values in the row into a validated object
  for (var index = 0; index < row.length; index++) {
    var skipField = false;

    if (!isComment) {
      if ((isUpdate && fields[index].type == fieldTypeEnums.id) || (isUpdate && fields[index].type == fieldTypeEnums.subId)) {
        //check for subTypeId
        if (fields[index].type == fieldTypeEnums.subId) {
          //parent values are not supposed to have subIds
          if ((row[index - 1] === "" || row[index - 1] == "-1") && row[index] === "") {
            missingSubId = true;
          }
        }
      }
      // first ignore entry that does not match the requirement specified in the fieldsToFilter
      else if (fieldsToFilter == FIELD_MANAGEMENT_ENUMS.standard && fields[index].isSubTypeField) {
        // skip the field
        skipField = true;
      } else if (fieldsToFilter == FIELD_MANAGEMENT_ENUMS.subType && !(fields[index].isSubTypeField || fields[index].isTypeAndSubTypeField)) {
        // skip the field
        skipField = true;
      }
      //comment/association fields are not part of the standard POST/PUT commands, so skip it
      else if (fields[index].isComment || fields[index].association) {
        var skipField = true;
      }
    }
    else {
      //In the comments mode, we just want the comment fields and the artifact ID field
      if (!fields[index].isComment && fields[index].type != fieldTypeEnums.id) {
        var skipField = true;
      }
    }

    // in all other cases add the field
    if (!skipField) {
      var value = null,
        customType = "",
        idFromName = 0;

      // double check data validation, convert dropdowns to required int values
      // sets both the value, and custom types - so that custom fields are handled correctly
      switch (fields[index].type) {

        // ID fields: restricted to numbers and blank on push, otherwise put
        case fieldTypeEnums.id:

          if (isUpdate && !isNaN(row[index])) {
            value = row[index];
          }
          customType = "IntegerValue";
          break;


        case fieldTypeEnums.subId:
          if (isUpdate && !isNaN(row[index]) && !isComment) {
            value = row[index];
          }
          customType = "IntegerValue";
          break;

        // INT fields
        case fieldTypeEnums.int:
          if (IS_GOOGLE) {
            if (row[index] != '' || row[index] == '0') {
              value = parseFloat(row[index]).toFixed(0);
            }
            else {
              value = '';
            }
            customType = "IntegerValue";
          } else {

            // only set the value if a number has been returned
            if (!isNaN(row[index])) {
              value = row[index];
              customType = "IntegerValue";
            }
          }
          break;
        case fieldTypeEnums.customInteger:
          if (IS_GOOGLE) {
            if (row[index] != '' || row[index] == '0') {
              value = parseFloat(row[index]).toFixed(0);
            }
            else {
              value = '';
            }
            customType = "IntegerValue";
          } else {

            // only set the value if a number has been returned
            if (!isNaN(row[index])) {
              value = row[index];
              customType = "IntegerValue";
            }
          }
          break;
        case fieldTypeEnums.customAutomationHost:
          if (IS_GOOGLE) {
            if (row[index] != '' || row[index] == '0') {
              value = parseFloat(row[index]).toFixed(0);
            }
            else {
              value = '';
            }
            customType = "IntegerValue";
          } else {

            // only set the value if a number has been returned
            if (!isNaN(row[index])) {
              value = row[index];
              customType = "IntegerValue";
            }
          }
          break;

        // DECIMAL fields
        case fieldTypeEnums.num:
          // only set the value if a number has been returned
          if (!isNaN(row[index])) {
            value = row[index];
            customType = "DecimalValue";
          }
          break;
        case fieldTypeEnums.customDecimal:
          // only set the value if a number has been returned
          if (!isNaN(row[index])) {
            value = row[index];
            customType = "DecimalValue";
          }
          break;

        // BOOL as Sheets has no bool validation, a yes/no dropdown is used
        case fieldTypeEnums.bool:
          // 'True' and 'False' don't work as dropdown choices, so have to convert back
          if (row[index] == "Yes") {
            value = "true";
            customType = "BooleanValue";
          } else if (row[index] == "No") {
            value = "false";
            customType = "BooleanValue";
          }
          break;
        case fieldTypeEnums.customBoolean:
          // 'True' and 'False' don't work as dropdown choices, so have to convert back
          if (row[index] == "Yes") {
            value = "true";
            customType = "BooleanValue";
          } else if (row[index] == "No") {
            value = "false";
            customType = "BooleanValue";
          }
          break;


        // DATES - parse the data and add prefix/suffix for WCF
        case fieldTypeEnums.date:
          if (row[index]) {
            if (IS_GOOGLE) {
              value = row[index];
            } else {
              // for Excel, dates are returned as days since 1900 - so we need to adjust this for JS date formats
              const dateInMs = (row[index] - DAYS_BETWEEN_1900_1970) * 86400 * 1000;
              value = convertLocalToUTC(new Date(dateInMs), dateInMs);
            }
            customType = "DateTimeValue";
          }
          break;
        case fieldTypeEnums.customDate:
          if (row[index]) {
            if (IS_GOOGLE) {
              value = row[index];
            } else {
              // for Excel, dates are returned as days since 1900 - so we need to adjust this for JS date formats
              const dateInMs = (row[index] - DAYS_BETWEEN_1900_1970) * 86400 * 1000;
              value = convertLocalToUTC(new Date(dateInMs), dateInMs);
            }
            customType = "DateTimeValue";
          }
          break;
        case fieldTypeEnums.customDateAndTime:
          if (row[index]) {
            if (IS_GOOGLE) {
              value = row[index];
            } else {
              // for Excel, dates are returned as days since 1900 - so we need to adjust this for JS date formats
              const dateInMs = (row[index] - DAYS_BETWEEN_1900_1970) * 86400 * 1000;
              value = convertLocalToUTC(new Date(dateInMs), dateInMs);
            }
            customType = "DateTimeValue";
          }
          break;
        // ARRAY fields are for multiselect lists - currently not supported so just push value into an array to make sure server handles it correctly
        case fieldTypeEnums.arr:
          if (row[index]) {
            value = [row[index]];
            customType = ""; // array fields not used for custom properties here
          }
          break;

        // DROPDOWNS - get id from relevant name, if one is present
        case fieldTypeEnums.drop:
          idFromName = getIdFromName(row[index], fields[index].values);
          if (idFromName) {
            value = idFromName;
            customType = "IntegerValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;
        case fieldTypeEnums.customList:
          idFromName = getIdFromName(row[index], fields[index].values);
          if (idFromName) {
            value = idFromName;
            customType = "IntegerValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;

        // MULTIDROPDOWNS - get id from relevant name, if one is present, set customtype to list value
        case fieldTypeEnums.multi:
          idFromName = getIdFromName(row[index], fields[index].values);
          if (idFromName) {
            value = [idFromName];
            customType = "IntegerListValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;
        case fieldTypeEnums.customMultiList:
          idFromName = getIdFromName(row[index], fields[index].values);
          if (idFromName) {
            value = [idFromName];
            customType = "IntegerListValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;

        // USER fields - get id from relevant name, if one is present
        case fieldTypeEnums.user:
          idFromName = getIdFromName(row[index], model.projectUsers);
          if (idFromName) {
            value = idFromName;
            customType = "IntegerValue";
          }
          else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;
        case fieldTypeEnums.customUser:
          idFromName = getIdFromName(row[index], model.projectUsers);
          if (idFromName) {
            value = idFromName;
            customType = "IntegerValue";
          }
          else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;

        // COMPONENT fields - get id from relevant name, if one is present
        case fieldTypeEnums.component:
          idFromName = getIdFromName(row[index], model.projectComponents);
          if (idFromName) {
            value = idFromName;
            // component is multi select for test cases but not for other artifacts
            customType = fields[index].isMulti ? "IntegerListValue" : "IntegerValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;

        // RELEASE fields - get id from relevant name, if one is present
        case fieldTypeEnums.release:
          idFromName = getIdFromName(row[index], model.projectActiveReleases);
          if (idFromName) {
            value = idFromName;
            customType = "IntegerValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;
        case fieldTypeEnums.customRelease:
          idFromName = getIdFromName(row[index], model.projectReleases);
          if (idFromName) {
            value = idFromName;
            customType = "IntegerValue";
          } else {
            if (row[index] != "") {
              entry.validationMessage = "Value not found for field '" + fields[index].name + "' . Please check the entry";
              return entry;
            }
          }
          break;

        // All other types
        default:
          // just assign the value to the cell - used for text
          value = row[index];
          customType = "StringValue";
          break;
      }


      // HIERARCHICAL ARTIFACTS:
      // handle hierarchy fields - if required: checks artifact type is hierarchical and if this field sets hierarchy
      if (artifactIsHierarchical && fields[index].setsHierarchy) {
        // first get the number of indent characters
        var indentInfo = countIndentCharacters(value, model.indentCharacter),
          indentCount = indentInfo.indentCount,
          trimCount = indentInfo.trimCount,
          indentPosition = setRelativePosition(indentCount, lastIndentPosition);
        // make sure to slice off the indent and spacing characters from the front
        value = value.slice(trimCount, value.length);

        // set the indent position for this row
        entry.indentPosition = indentPosition;
      }

      // CUSTOM FIELDS:
      // check whether field is marked as a custom field and as the required property number
      if (fields[index].isCustom && fields[index].propertyNumber) {

        // if field has data create the object
        if (value) {
          var customObject = {};
          customObject.PropertyNumber = fields[index].propertyNumber;
          customObject[customType] = value;

          entry.CustomProperties.push(customObject);
        }

        // STANDARD FIELDS:
        // add standard fields in standard way - only add if field contains data
      } else if (value) {
        if (isComment) {

          if (fields[index].type == fieldTypeEnums.id) {
            //for comments, we always have the ArtifactID
            entry.ArtifactId = value;
          }
          else {//add the value normally
            entry[fields[index].field] = (customType == "IntegerListValue") ? [value] : value;
          }
        }
        else {
          // if the standard field is a multi select type as set in the switch above, pass the value through in an array
          entry[fields[index].field] = (customType == "IntegerListValue") ? [value] : value;
        }
      }
    }

  }
  //check for special case, when we have a new artifact on an update sheet
  if (missingSubId && model.currentArtifact.allowsCreateOnUpdate) {
    //set this flag
    entry.createOnUpdate = true;
  }
  //double checking if the comment object has at least the necessary fields (i.e.: comment + ID). If not, add the correspondent fields
  if (isComment) {
    if (entry.ArtifactId && entry.Text) {
      return entry
    }
    else if (entry.Text && sourceId) {
      //we still need to get the artifactID
      entry.ArtifactId = sourceId;
      return entry
    }
    else return false;
  }
  else {
    return entry;
  }
}

//Converts a local time to UTC time
function convertLocalToUTC(convertedDate, originalDate) {
  originalDate = new Date(originalDate).toUTCString();
  var d = new Date();
  var offsetMinutes = d.getTimezoneOffset();
  var utcDate = new Date(convertedDate.getTime() + offsetMinutes * 60000);
  return utcDate.toISOString();
}

//Converts a UTC to local
function convertUTCtoLocal(originalDate) {
  var d = new Date();
  var offsetMinutes = d.getTimezoneOffset();
  var utcDate = new Date(originalDate.getTime() + offsetMinutes * 60000);
  return utcDate.toISOString();
}

// function creates a correctly formatted artifact object ready to send to Spira
// it works through each field type to validate and parse the values so object is in correct form
// any field that does not pass validation receives a null value
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: model - full model with info about fields, dropdowns, users, etc
// @param: fieldTypeEnums - object of all field types with enums
// @param: originId - the ID of the artifact that originated this association
function createAssociationEntryFromRow(row, model, fieldTypeEnums, originId) {

  var fields = model.fields,
    finalEntry = [],
    sourceId, sourceTypeId,
    destTypeId = [],
    destId = [];

  //Starting get the data from the row
  sourceTypeId = model.currentArtifact.id;

  //get all the data we need from each row
  for (var index = 0; index < row.length; index++) {

    destTypeId = [];
    destId = [];
    if (fields[index].type == fieldTypeEnums.id) {
      sourceId = row[index]
    }

    if (row[index] && fields[index].association) {
      var associationType = fields[index].association;

      //depending on the user's language, the console can misinterpret commas as dots - this fixes it
      var associationText = (row[index] + '').replace('.', ',');
      //removing spaces from the inputs, since it can cause errors
      associationText = associationText.replace(/\s/g, '');
      var associationIds = (associationText).split(',');
      //work every chunk of data
      associationIds.forEach(function (item) {
        if (associationType == params.associationEnums.req2req) { destTypeId.push(params.artifactEnums.requirements); }
        destId.push(item);
      });

      //build the entry object (return object)
      destId.forEach(function entryBuilder(item, pos) {
        var singleEntry = {};
        if (sourceId) {
          if (associationType == params.associationEnums.req2req) { singleEntry.SourceArtifactId = sourceId; }
          if (associationType == params.associationEnums.tc2req || associationType == params.associationEnums.tc2rel) { singleEntry.TestCaseId = sourceId; }
          if (associationType == params.associationEnums.tc2ts) { singleEntry.TestSetTestCaseId = sourceId; }
        }
        else if (originId) { //if we just created this artifact, its ID is not in the spreadsheet yet
          if (associationType == params.associationEnums.req2req) { singleEntry.SourceArtifactId = originId; }
          if (associationType == params.associationEnums.tc2req || associationType == params.associationEnums.tc2rel) { singleEntry.TestCaseId = originId; }
          if (associationType == params.associationEnums.tc2ts) { singleEntry.TestSetTestCaseId = originId; }
        }
        if (associationType == params.associationEnums.req2req) {
          singleEntry.SourceArtifactTypeId = sourceTypeId;
          singleEntry.DestArtifactId = Number(item);
          singleEntry.DestArtifactTypeId = destTypeId[pos];
          //association is always "related to"
          singleEntry.ArtifactLinkTypeId = Number("1");
        } else if (associationType == params.associationEnums.tc2req) { singleEntry.RequirementId = Number(item); }
        else if (associationType == params.associationEnums.tc2rel) { singleEntry.ReleaseId = Number(item); }
        else if (associationType == params.associationEnums.tc2ts) { singleEntry.TestSetId = Number(item); }

        //append to the export array
        finalEntry.push(singleEntry);
      });
    }
  }
  return finalEntry;
}

// find the corresponding ID for a string value - eg from a dropdown
// dropdowns can only contain one item per row so we have to now get the IDs for sending to Spira
// @param: string - the string of the name value specified
// @param: list - the array of items with keys for id and name values
function getIdFromName(string, list) {
  for (var i = 0; i < list.length; i++) {
    if (setListItemDisplayName(list[i]) == string) {
      return list[i].id;

      // if there's no match with the item, let's try and match on just the name part of the list item
      // this code is included to accomodate users who create their spreadsheets elsewhere and then dump the data in here without knowing the ids
    } else if (compareItemName(string, list[i].name)) {
      return list[i].id;
    }
    /*else{
        //the provided name is not part of the list - use a value that indicates error
        return -1;
    }*/
  }
  // return 0 if there's no match from either method
  return 0;
}

// for dropdown items we need to use the id as well as the name to make sure the entries are unique - so return a standard format here
// @param: item - object of the list item - contains a name and id
// returns the correctly formatted string - so that it is always set consistently
function setListItemDisplayName(item) {

  return item.name + " (#" + item.id + ")";

}

// removes spaces and upper case from strings to compare them
// @param: string - the base string to compare (user input)
// @param: value - the value to be compared with (system input)
// returns true if there's a match or false if there's not
function compareItemName(string, value) {

  var newString = string.toLowerCase().trim();
  var newValue = value.toLowerCase().trim();

  if (newString == newValue) {
    return true;
  }
  else {
    return false;
  }

}


// finds and returns the field name for the specific artifiact's ID field
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldTypeEnums - object of all field types with enums
// @param: getSubType - optioanl bool to specify to return the subtype Id field, not the normal field (where two exist)
// @param: artifactId - the artifact type Id that was sent to Spira
function getIdFieldName(fields, fieldTypeEnums, getSubType) {
  var fieldToLookup = "";
  for (var i = 0; i < fields.length; i++) {
    var fieldToLookup = getSubType ? "subId" : "id";
    if (fields[i].type == fieldTypeEnums[fieldToLookup]) {
      return fields[i].field;
    }
  }
  return null;
}


// returns the count of the number of indent characters and returns the value
// @param: field - a single field string - one already designated as containing hierarchy info
// @param: indentCharacter - the character used to denote an indent - e.g. ">"
function countIndentCharacters(field, indentCharacter) {
  var indentCount = 0,
    trimCount = 0;
  //check for field value and indent character
  if (field && field[0] === indentCharacter) {
    //increment indent counter while there are '>'s present
    while (field[0] === indentCharacter || field[0] === " ") {
      // in all cases as to the trim count
      trimCount++;
      // add to the indent count if we have an indent character
      if (field[0] === indentCharacter) {
        indentCount++;
      }

      //get entry length for slice
      var len = field.length;
      //slice the first character off of the entry
      field = field.slice(1, len);
    }
  }
  return {
    indentCount: indentCount,
    trimCount: trimCount
  };
}


// returns the correct relative indent position - based on the previous relative indent and other logic (int neg, pos, or zero)
// Currently the API does not support a call to place an artifact at a certain location.
// @param: indentCount - int of the number of indent characters set by user
// @param: lastIndentPosition - int sum of the actual indent positions used for the preceding entries
function setRelativePosition(indentCount, lastIndentPosition) {
  // the first time this is called, last position will be null
  if (lastIndentPosition === null) {
    return 0;
  } else if (indentCount > lastIndentPosition) {
    // only indent one level at a time
    return lastIndentPosition + 1;
  } else {
    // this will manage indents of same level or where outdents are required
    return indentCount;
  }
}

// anaylses the response from posting an item to Spira, and handles updating the log and displaying any messages to the user
function processSendToSpiraResponse(i, sentToSpira, entriesForExport, artifact, log, isComment, isAssociation) {
  var response = {};
  response.details = sentToSpira;
  var operationString = "";
  if (isComment) { operationString = " Comment "; }
  if (isAssociation) { operationString = " Association "; }

  // handle success and error cases
  if (sentToSpira.error) {
    log.errorCount++;
    response.error = true;
    var errorCode = 0;
    if (IS_GOOGLE) {
      errorCode = response.details.httpCode;
    }
    else {
      errorCode = sentToSpira.errorMessage.status;
    }




    //handling different error messages
    if (errorCode == 409) {
      response.message = "Concurrency Date conflict: please reload your data and try again.";
    }
    else {
      if (errorCode == 400) {
        //Handling different types of error description from Spira API
        if (IS_GOOGLE) {
          try {

            var serverReply = sentToSpira.errorMessage;
            var regExp = new RegExp(/<Message>(.*?)<\/Message>/i).exec(serverReply);;
            response.message = regExp[regExp.length - 1];
          }
          catch (err) {
            try {
              var serverReply2 = sentToSpira.errorMessage;
              var regExp2 = new RegExp(/<string [^>]+>(.*?)<\/string>/i).exec(serverReply2);
              response.message = regExp2[regExp2.length - 1];
            }
            catch (err2) {
              try {
                var serverReply3 = sentToSpira.errorMessage;
                var regExp3 = new RegExp(/'[^']*'/i).exec(serverReply3);
                response.message = regExp3[regExp3.length - 1];
              }
              catch (err3) {
                response.message = "Update attempt failed. Please check your data.";
              }
            }
          }
        }
        else {
          try {
            let parser = new DOMParser();
            let doc = parser.parseFromString(sentToSpira.errorMessage.response.text, 'text/html');
            try {
              let errorMessage2 = doc
                .querySelector('p:nth-of-type(2)')
                .innerHTML.split(" '")[1]
                .split("'.")[0];
              response.message = errorMessage2;
            }
            catch (err) {
              try {
                let htmlError = doc.querySelector('string').innerHTML;
                response.message = htmlError;
              }
              catch (err) {
                try {
                  let htmlError2 = Array.from(doc.querySelectorAll('messages message'));
                  let completeMessage = '';
                  htmlError2.forEach((value) => {
                    completeMessage += value.innerHTML + ' / ';
                  });
                  response.message = completeMessage.slice(0, -2);
                }
                catch (err) {
                  response.message = "Update attempt failed. Please check your data or contact your system administrator.";
                }
              }
            }
          }
          catch (err) {
            response.message = "Update attempt failed. Please check your data or contact your system administrator.";
          }
        }
      }
      else {
        response.message = sentToSpira.errorMessage;
      }
    }
    //Sets error HTML modals
    if (artifact.hierarchical) {
      // if there is an error on any hierarchical artifact row, break out of the loop to prevent entries being attached to wrong parent
      popupShow('Error sending ' + operationString + (i + 1) + ' of ' + (entriesForExport.length) + ' - sending stopped to avoid indenting entries incorrectly', 'Progress')
    } else {
      popupShow('Error sending ' + operationString + (i + 1) + ' of ' + (entriesForExport.length), 'Progress');
    }
  }
  else {
    log.successCount++;
    response.newId = sentToSpira.newId;

    // if artifact is hierarchical save relevant information to work out how to indent
    if (artifact.hierarchical) {
      response.details.hierarchyInfo = {
        id: sentToSpira.newId,
        indent: entriesForExport[i].indentPosition
      }
    }
    //modal that displays the status of each artifact sent
    popupShow('Sent ' + operationString + (i + 1) + ' of ' + (entriesForExport.length) + '...', 'Progress');
  }

  // finally write out the response to the log and return
  log.entries.push(response);
  return log;
}









/*
 * ==================
 * GETTING FROM SPIRA
 * ==================
 *
 * get all items of an artifact and place contents in the sheet
 *
 */

// GOOGLE SPECIFIC VARIATION OF THIS FUNCTION handles getting paginated artifacts from Spira and saving them as a single array
// @param: model - full model object from client
// @param: fieldTypeEnums - enum of fieldTypes used
function getFromSpiraGoogle(model, fieldTypeEnums, advancedMode) {
  var requiredSheetName;
  var singleArtifactId = null;

  if (model.currentOperation) {
    //administrator mode

    var operation = model.operations.filter(function (operation) {
      return operation.id == model.currentOperation;
    })[0];

    if (operation.type == "send-system") {
      //system wide operations
      requiredSheetName = model.currentArtifact.name + ", system";

    }
    else if (operation.type == "send-template" || operation.type == "get-template") {
      //template-based operations
      requiredSheetName = model.currentArtifact.name + ", TP-" + model.currentTemplate.id;
    }
  }
  else {
    //stardard artifact functions
    requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  if (sheet.getName() != requiredSheetName) {
    return operationComplete(STATUS_ENUM.wrongSheet, false);
  }

  resetSheetColors(model, fieldTypeEnums, sheet);

  //0. If we are not using advanced mode, we need to make sure the model excludes the non-advanced fields

  var fields = model.fields;
  if (!advancedMode) {
    model.fields = fields.filter(function (item, index) {
      if (!item.isAdvanced) {
        return item;
      }
    })
  }

  // 1. get from spira
  // note we don't do this by getting the count of each artifact first, because of a bug in getting the release count
  var currentPage = 0;
  var artifacts = [];
  var getNextPage = true;

  //if this artifact supports getting a single result, handle it
  if (model.currentArtifact.allowGetSingle) {
    // Custom Lists
    if (model.currentArtifact.id == params.artifactEnums.customLists) {
      singleArtifactId = model.currentList.id;
    }
  }

  while (getNextPage && currentPage < 100) {
    var startRow = (currentPage * GET_PAGINATION_SIZE) + 1;
    var pageOfArtifacts = getArtifacts(
      model.user,
      model.currentProject.id,
      model.currentArtifact.id,
      startRow,
      GET_PAGINATION_SIZE,
      singleArtifactId,
      model.currentTemplate.id
    );
    // if we got a non empty array back then we have artifacts to process
    if (pageOfArtifacts.length || pageOfArtifacts.Values) {
      artifacts = artifacts.concat(pageOfArtifacts);
      //handling special allowGetSingle cases
      if (model.currentArtifact.allowGetSingle && singleArtifactId > 0) {
        //we just want one artifact - one query it's enough
        getNextPage = false;
      }
      // if we got less artifacts than the max we asked for, then we reached the end of the list in this request - and should stop
      if (pageOfArtifacts.length < GET_PAGINATION_SIZE) {
        getNextPage = false;
        // if we got the full page size back then there may be more artifacts to get
      } else {
        currentPage++;
      }
      // if we got no artifacts back, stop now
    } else {
      getNextPage = false;
    }
  }

  // 2. if there were no artifacts at all break out now
  if (!artifacts.length) return "no artifacts were returned";

  // 3. Make sure hierarchical artifacts are ordered correctly
  if (model.currentArtifact.hierarchical) {
    artifacts.sort(function (a, b) {
      return a.indentLevel < b.indentLevel ? -1 : 1;
    });
  }

  // 4. if artifact has subtype that needs to be retrieved separately, do so
  if (model.currentArtifact.hasSubType) {
    // find the id field
    var idFieldNameArray = model.fields.filter(function (field) {
      return field.type === fieldTypeEnums.id;
    });
    // if we have an id field, then we can find the id number for each artifact in the array
    if (idFieldNameArray && idFieldNameArray[0].field) {
      var idFieldName = idFieldNameArray[0].field;
      var artifactsWithSubTypes = [];
      artifacts.forEach(function (art) {
        artifactsWithSubTypes.push(art);
        var subTypeArtifacts = getArtifacts(
          model.user,
          model.currentProject.id,
          model.currentArtifact.subTypeId,
          null,
          null,
          art[idFieldName],
          model.currentTemplate.id
        );
        // take action if we got any sub types back - ie if they exist for the specific artifact
        if (subTypeArtifacts && (subTypeArtifacts.length || subTypeArtifacts.Values.length)) {

          if (subTypeArtifacts.Values) {
            //some subArtifacts, such as Custom Values, require this adjustment
            subTypeArtifacts = subTypeArtifacts.Values;
          }

          var subTypeArtifactsWithMeta = subTypeArtifacts.map(function (sub) {
            sub.isSubType = true;
            sub.parentId = art[idFieldName];
            return sub;
          });
          // now add the steps into the original object
          artifactsWithSubTypes = artifactsWithSubTypes.concat(subTypeArtifactsWithMeta);
        }
      })
      // update the original array (I know that mutation is bad, but it makes things easy here)
      artifacts = artifactsWithSubTypes;
    }
  }

  // 5. create 2d array from data to put into sheet
  var artifactsAsCells = matchArtifactsToFields(
    artifacts,
    model.currentArtifact,
    model.fields,
    fieldTypeEnums,
    model.projectUsers,
    model.projectComponents,
    model.projectActiveReleases,
    model.projectReleases
  );

  // 6. add data to sheet
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = spreadSheet.getActiveSheet(),
    range = sheet.getRange(2, 1, artifacts.length, model.fields.length);

  range.setValues(artifactsAsCells);

  return JSON.parse(JSON.stringify(artifactsAsCells));
}

// EXCEL SPECIFIC VARIATION OF THIS FUNCTION handles getting paginated artifacts from Spira and displaying them in the UI
// @param: model: full model object from client
// @param: enum of fieldTypeEnums used
function getFromSpiraExcel(model, fieldTypeEnums) {
  return Excel.run(function (context) {
    var fields = model.fields;
    var sheet = context.workbook.worksheets.getActiveWorksheet(),
      sheetRange = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS_PER_PAGE, fields.length);
    sheet.load("name");
    sheetRange.load("values");
    var requiredSheetName;

    if (model.currentOperation) {
      //administrator mode

      var operation = model.operations.filter(function (operation) {
        return operation.id == model.currentOperation;
      })[0];

      if (operation.type == "send-system") {
        //system wide operations
        requiredSheetName = model.currentArtifact.name + ", system";

      }
      else if (operation.type == "send-template" || operation.type == "get-template") {
        //template-based operations
        requiredSheetName = model.currentArtifact.name + ", TP-" + model.currentTemplate.id;
      }
    }
    else {
      //stardard artifact functions
      requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
    }

    return context.sync()
      .then(function () {
        // only get the data if we are on the right sheet - the one with the template loaded on it
        if (sheet.name == requiredSheetName) {
          //first, reset sheet rows (if we had data from the last run)
          resetSheet(model);
          //then, clear the background colors of the spreadsheet (in case we had any errors in the last run)
          resetSheetColors(model, fieldTypeEnums, sheetRange);
          dataBaseValidationSetter(requiredSheetName, model, fieldTypeEnums, context);
          return getDataFromSpiraExcel(model, fieldTypeEnums).then((response) => {
            //error handling           
            if (typeof(response) === 'string' && (response.indexOf('noData') !== -1 || response.indexOf('no artifacts')!== -1)) {
              return operationComplete(STATUS_ENUM.noData, false);
            }
            else {
              return processDataFromSpiraExcel(response, model, fieldTypeEnums)
            }
          });
        } else {
          return operationComplete(STATUS_ENUM.wrongSheet, false);
        }
      })
  })
}

// resets validations - used before a GET command
// @param: sheetName - current sheet name
function resetSheet(model) {

  Excel.run(function (ctx) {
    var fields = model.fields;
    //complete data range from old data
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS_PER_PAGE, fields.length);
    range.delete(Excel.DeleteShiftDirection.up);

    ctx.sync();

    var dataBaseSheetName = createDatabaseSheetName(params.dataSheetName, model.currentProject.id, model.currentArtifact.id);

    //clear database worksheet
    var worksheet = context.workbook.worksheets.getItemOrNullObject(dataBaseSheetName);
    worksheet.getRange().clear();

    return ctx.sync();
  }).catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
    }
  });
}


// EXCEL SPECIFIC VARIATION OF THIS FUNCTION handles getting paginated artifacts from Spira and saving them as a single array
// @param: model - full model object from client
// @param: fieldTypeEnums - enum of fieldTypes used
async function getDataFromSpiraExcel(model, fieldTypeEnums) {
  // 1. get from spira
  // note we don't do this by getting the count of each artifact first, because of a bug in getting the release count
  var currentPage = (EXCEL_MAX_ROWS_PER_PAGE / GET_PAGINATION_SIZE) * (model.selectedPage - 1);
  var results = {};
  results.firstRecord = (currentPage * GET_PAGINATION_SIZE) + 1;
  var artifacts = [];
  var getNextPage = true;
  var singleArtifactId = null;

  //if this artifact supports getting a single result, handle it
  if (model.currentArtifact.allowGetSingle) {
    // Custom Lists
    if (model.currentArtifact.id == params.artifactEnums.customLists) {
      singleArtifactId = model.currentList.id;
    }
  }

  async function getArtifactsPage(startRow) {
    await getArtifacts(
      model.user,
      model.currentProject.id,
      model.currentArtifact.id,
      startRow,
      GET_PAGINATION_SIZE,
      singleArtifactId,
      model.currentTemplate.id
    ).then(function (response) {
      // if we got a non empty array back then we have artifacts to process
      if (response.body && (response.body.length || response.body.Values)) {
        artifacts = artifacts.concat(response.body);
        //handling special allowGetSingle cases
        if (model.currentArtifact.allowGetSingle && singleArtifactId > 0) {
          //we just want one artifact - one query it's enough
          getNextPage = false;
        }
        // if we got less artifacts than the max we asked for, then we reached the end of the list in this request - and should stop
        if (response.body && response.body.length < GET_PAGINATION_SIZE) {
          getNextPage = false;
          // if we got the full page size back then there may be more artifacts to get
        } else {
          currentPage++;
        }
        //if we got more artfacts than the maximum global variable, we should stop
        if (startRow >= ((EXCEL_MAX_ROWS_PER_PAGE * model.selectedPage) - GET_PAGINATION_SIZE)) {
          getNextPage = false;
        }
        // if we got no artifacts back, stop now
      } else {
        getNextPage = false;
      }
    })
      .catch(/*fail quietly*/);
  }


  while (getNextPage && currentPage < ARTIFACT_MAX_PAGES) {
    var startRow = (currentPage * GET_PAGINATION_SIZE) + 1;
    //update the log every time - to capture the last one
    results.lastRecord = startRow + GET_PAGINATION_SIZE - 1;
    await getArtifactsPage(startRow);
  }

  // 2. if there were no artifacts at all break out now
  if (!artifacts.length) return "no artifacts were returned";

  // 3. Make sure hierarchical artifacts are ordered correctly
  if (model.currentArtifact.hierarchical) {
    artifacts.sort(function (a, b) {
      return a.indentLevel < b.indentLevel ? -1 : 1;
    });
  }
  // 4. if artifact has subtype that needs to be retrieved separately, do so
  if (model.currentArtifact.hasSubType) {
    // find the id field
    var idFieldNameArray = model.fields.filter(function (field) {
      return field.type === fieldTypeEnums.id;
    });


    // if we have an id field, then we can find the id number for each artifact in the array
    if (idFieldNameArray && idFieldNameArray[0].field) {
      //function called below in the foreach call
      async function getArtifactSubs(art) {
        await getArtifacts(
          model.user,
          model.currentProject.id,
          model.currentArtifact.subTypeId,
          null,
          null,
          art[idFieldName],
          model.currentTemplate.id
        ).then(function (response) {
          try {
            // take action if we got any sub types back - ie if they exist for the specific artifact
            if (response.body && (response.body.length || response.body.Values.length)) {
              if (response.body.Values) {
                //some subArtifacts, such as Custom Values, require this adjustment
                response.body = response.body.Values;
              }
              var subTypeArtifactsWithMeta = response.body.map(function (sub) {
                sub.isSubType = true;
                sub.parentId = art[idFieldName];
                return sub;
              });
              // now add the steps into the original object
              artifactsWithSubTypes = artifactsWithSubTypes.concat(subTypeArtifactsWithMeta);
            }
          }
          catch (err) {
            //do nothing - just proceed to the next artifact
          }
        })
      };

      var idFieldName = idFieldNameArray[0].field;
      var artifactsWithSubTypes = [];

      for (var i = 0; i < artifacts.length; i++) {
        artifactsWithSubTypes.push(artifacts[i]);
        await getArtifactSubs(artifacts[i]);
      }
      // update the original array (I know that mutation is bad, but it makes things easy here)
      artifacts = artifactsWithSubTypes;
    }
  }
  results.artifacts = artifacts;
  return results;
}

// EXCEL SPECIFIC to process all the data retrieved from Spira and then display it
// @param: artifacts: array of raw data from Spira (with subtypes already present if needed)
// @param: model: full model object from client
// @param: enum object of the different fieldTypeEnums
function processDataFromSpiraExcel(log, model, fieldTypeEnums) {
  //pre-handling the log info
  var results = {};
  results.firstRecord = log.firstRecord;
  results.lastRecord = log.lastRecord;

  var artifacts = log.artifacts;

  // 5. create 2d array from data to put into sheet
  var artifactsAsCells = matchArtifactsToFields(
    artifacts,
    model.currentArtifact,
    model.fields,
    fieldTypeEnums,
    model.projectUsers,
    model.projectComponents,
    model.projectActiveReleases,
    model.projectReleases
  );

  // 6. add data to sheet
  return Excel.run({ delayForCellEdit: true }, function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet(),
      range = sheet.getRangeByIndexes(1, 0, artifacts.length, model.fields.length);
    range.values = artifactsAsCells;

    return context.sync()
      .then(function () {

        results.artifacts = artifactsAsCells;

        return results;
      })
  })
}

// matches data against the fields to be shown in the spreadsheet - not all data fields are shown
// @param: artifacts - array of the artifact objects we GOT from Spira
// @param: artifactMeta - object of the meta information about the artifact
// @param: fields - array of the fields that make up the sheet display
// @param: fieldTypeEnums - enum object of the different fieldTypes
// @param: users - array of the user objects
// @param: components - array of the component objects
// @param: releases - array of the release objects
function matchArtifactsToFields(artifacts, artifactMeta, fields, fieldTypeEnums, users, components, activeReleases, releases) {
  return artifacts.map(function (art) {
    return fields.map(function (field) {
      var originalFieldValue = "";

      // handle custom fields
      if (field.isCustom) {
        //main artifacts field
        if (!field.isSubTypeField && !art.isSubType) {
          // if we have any custom props
          if (art.CustomProperties && art.CustomProperties.length) {
            // look for a match for the current field
            var customProp = art.CustomProperties.filter(function (custom) {
              return custom.Definition.CustomPropertyFieldName == field.field;
            });
            // if the property exists and isn't null - do a null check to handle booleans properly
            if (typeof customProp != "undefined" && customProp.length && customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]] !== null) {
              originalFieldValue = customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]];
            }
          }
        } else if (field.isSubTypeField && art.isSubType) {
          // subtypes custom fields
          if (art.CustomProperties && art.CustomProperties.length) {
            // look for a match for the current field
            var customProp = art.CustomProperties.filter(function (custom) {
              return custom.Definition.CustomPropertyFieldName == field.field;
            });
            // if the property exists and isn't null - do a null check to handle booleans properly
            if (typeof customProp != "undefined" && customProp.length && customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]] !== null) {
              originalFieldValue = customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]];
            }
          }
        }

      }
      // handle subtype fields. If dual type, make sure this is a subtype (avoid main types entering here)
      else if (field.isSubTypeField || (field.isTypeAndSubTypeField && art.isSubType)) {
        if (artifactMeta.hasSubType && art.isSubType) {
          // first check to make sure the field exists in the artifact data
          if (typeof art[field.field] != "undefined" && art[field.field]) {
            originalFieldValue = art[field.field]
          }
        }
        // handle standard fields
      } else if (!art.isSubType) {
        // first check to make sure the field exists in the artifact data
        if (typeof art[field.field] != "undefined" && art[field.field] != null && art[field.field]) {
          originalFieldValue = art[field.field];
        }
      }
      // handle list values - turn from the id to the actual string so the string can be displayed
      if (
        field.type == fieldTypeEnums.drop ||
        field.type == fieldTypeEnums.customList ||
        field.type == fieldTypeEnums.multi ||
        field.type == fieldTypeEnums.customMultiList ||
        field.type == fieldTypeEnums.user ||
        field.type == fieldTypeEnums.component ||
        field.type == fieldTypeEnums.release ||
        field.type == fieldTypeEnums.customUser ||
        field.type == fieldTypeEnums.customRelease
      ) {
        // a field can have display overrides - if one of these overrides is in the artifact field specified, then this is returned instead of the lookup - used specifically to make sure RQ Epics show as Epics 
        if (field.displayOverride && field.displayOverride.field && field.displayOverride.values && field.displayOverride.values.includes(art[field.displayOverride.field])) {
          var shoudOverride = true;
          field.values.forEach(function (value) {
            if (field.displayOverride.values.includes(value.name)) {
              shoudOverride = false;
            }
          });
          if (shoudOverride) {
            if (IS_GOOGLE) {
              return art[field.displayOverride.field];
            }
            else {
              return art[field.displayOverride.field];
            }
          }
          else {
            //this is an override field, but we don't want it to be - handle as any other
            var fieldValueForLookup = Array.isArray(originalFieldValue) ? originalFieldValue[0] : originalFieldValue;
            var fieldName = getListValueFromId(
              fieldValueForLookup,
              field.type,
              fieldTypeEnums,
              field.values,
              users,
              components,
              activeReleases,
              releases
            );

            if (fieldName && fieldValueForLookup) {
              fieldName = fieldName + " (#" + fieldValueForLookup + ")";
            }

            return fieldName;
          }
        } else {
          // handle multilist fields (custom props or components for some artifacts) - we can only display one in Excel so pick the first in the array to match
          var fieldValueForLookup = Array.isArray(originalFieldValue) ? originalFieldValue[0] : originalFieldValue;
          var fieldName = getListValueFromId(
            fieldValueForLookup,
            field.type,
            fieldTypeEnums,
            field.values,
            users,
            components,
            activeReleases,
            releases
          );

          if (fieldName && fieldValueForLookup) {
            fieldName = fieldName + " (#" + fieldValueForLookup + ")";
          }

          return fieldName;

        }

        // handle date fields 
      } else if (field.type == fieldTypeEnums.date || field.type == fieldTypeEnums.customDate || field.type == fieldTypeEnums.customDateAndTime) {
        if (IS_GOOGLE) {
          var dateValue = new Date(originalFieldValue);
          if (isNaN(dateValue.getTime())) { //invalid date
            return null;
          }
          else {
            return dateValue;
          }
        } else {
          if (originalFieldValue) {
            var jsObj = new Date(originalFieldValue);
            return JSDateToExcelDate(jsObj);
          } else {
            return "";
          }
        }

        // handle booleans - need to make sure null values are ignored ie treated differently to false
      } else if (field.type == fieldTypeEnums.bool || field.type == fieldTypeEnums.customBoolean) {

        /*in case of null/blank values, the return value depends on the characteristic of the artifact (i.e.: we want to see an explicit "No" when
        that field is shared between a main and a subtypes)*/
        if (artifactMeta.hasDualValues) {
          return originalFieldValue ? "Yes" : originalFieldValue === false ? "No" : "No";
        }
        else {
          return originalFieldValue ? "Yes" : originalFieldValue === false ? "No" : "";
        }
        // handle hierarchical artifacts
      } else if (field.setsHierarchy) {
        return makeHierarchical(originalFieldValue, art.IndentLevel);
        // handle artifacts that have extra information we can display to the user - ie where there is a linked test step this will add the information about the link at the end of the field
      } else if (field.extraDataField && art[field.extraDataField]) {
        return `${originalFieldValue} ${field.extraDataPrefix ? field.extraDataPrefix + ":" : ""}${art[field.extraDataField]}`;
      } else {
        return originalFieldValue;

      }
    });
  })
}

// takes an id for a lookup field and returns the string to display
// @param: id - int of the id to lookup
// @param: type - enum of the type of filed we need to look up
// @param: fieldTypeEnums - enum object of the different fieldTypes
// @param: fieldValues - array of the value objects for bespoke lookups
// @param: users - array of the user objects
// @param: components - array of the component objects
// @param: releases - array of the release objects
function getListValueFromId(id, type, fieldTypeEnums, fieldValues, users, components, activeReleases, releases) {
  var match = null;
  switch (type) {
    case fieldTypeEnums.drop:
    case fieldTypeEnums.customList:
    case fieldTypeEnums.customMultiList:
    case fieldTypeEnums.multi:
      match = fieldValues.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.user:
      match = users.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.customUser:
      match = users.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.component:
      match = components.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.release:
      match = activeReleases.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.customRelease:
      match = releases.filter(function (val) { return val.id == id; });
      break;
  }
  return typeof match != "undefined" && match && match.length ? match[0].name : "";
}

function makeHierarchical(value, indent) {
  var indentIncrements = Math.floor(indent.length / 3) - 1;
  var indentText = "";
  for (var i = 0; i < indentIncrements; i++) {
    indentText += "> ";
  }
  indentText += value;
  return indentText;
}

//Convert JS date object to excel format - Excel dates start at 1900 not 1970
//@param: inDate - js date object
function JSDateToExcelDate(inDate) {
  var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
  return returnDateTime.toString().substr(0, 20);
}