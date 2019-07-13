/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 *
 */
export { 
    clearAll, 
    getProjects, 
    getFromSpiraExcel, 
    warn, 
    sendToSpira, 
    operationComplete,
    getComponents,
    getCustoms,
    getBespoke,
    getReleases,
    getUsers,
    templateLoader,
    error
};

import { showPanel, hidePanel } from './taskpane.js';

/*
 * =======
 * TODO
 * =======
 
 - make sure when you change project / art the get / send buttons are disabled
 - check what happens when add more rows from get than are on sheet. Does the validation get copied down?
 - do we need PUT? Existing customers want it but it causes so much hassle and misuse
 - TODO: disable "send to spira" after have done a get
 - better handling of trying to do a put
 - try using it for several times in a row and fix any bugs
 */




// globals
var API_BASE = '/services/v5_0/RestService.svc/projects/',
	API_BASE_NO_SLASH = '/services/v5_0/RestService.svc/projects',
	ART_ENUMS = {
		requirements: 1,
		testCases: 2,
		incidents: 3,
		releases: 4,
		testRuns: 5,
		tasks: 6,
		testSteps: 7,
		testSets: 8
	},
	INITIAL_HIERARCHY_OUTDENT = -20,
	GET_PAGINATION_SIZE = 100,
	EXCEL_MAX_ROWS = 1000,
	FIELD_MANAGEMENT_ENUMS = {
		all: 1,
		standard: 2,
		subType: 3
	},
	STATUS_ENUM = {
		allSuccess: 1,
		allError: 2,
		someError: 3,
		wrongSheet: 4
	},
	STATUS_MESSAGE_GOOGLE = {
		1: "All done! To send more data over, clear the sheet first.",
		2: "Sorry, but there were some problems (these cells are marked in red). Check the notes on the relevant ID field for explanations.",
		3: "We're really sorry, but we couldn't send anything to SpiraTeam - please check notes on the ID fields for more information.",
		4: "You are not on the correct worksheet. Please go to the sheet that matches the one listed on the Spira taskpane."
	},
	STATUS_MESSAGE_EXCEL = {
		1: "All done! To send more data over, clear the sheet first.",
		2: "Sorry, but there were some problems (these cells are marked in red). Check the cells at the end of relevant rows for explanations.",
		3: "We're really sorry, but we couldn't send anything to SpiraTeam - please check cells at the end of the row for more information.",
		4: "You are not on the correct worksheet. Please go to the sheet that matches the one listed on the Spira taskpane."
	},
    CUSTOM_PROP_TYPE_ENUM = {
      1: "StringValue",
      2: "IntegerValue",
      3: "DecimalValue",
      4: "BooleanValue",
      5: "DateTimeValue",
      6: "IntegerValue",
      7: "IntegerListValue",
      8: "IntegerValue"
      
    },
  INLINE_STYLING = "style='font-family: sans-serif'",
  IS_GOOGLE = typeof UrlFetchApp != "undefined";

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
	.setTitle('SpiraTeam by Inflectra');

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



//clears active sheet in spreadsheet
function clearAll() {

  if (IS_GOOGLE) {
    // get active spreadsheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = spreadSheet.getActiveSheet(),
		lastColumn = sheet.getMaxColumns(),
		lastRow = sheet.getMaxRows();

    sheet.clear()
  
    // clears data validations and notes from the entire sheet
    var range = sheet.getRange(1, 1, lastRow, lastColumn);
    range.clearDataValidations().clearNote();
  
    // remove any protections on the sheet
    var protections = spreadSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
       var protection = protections[i];
       if (protection.canEdit()) {
         protection.remove();
       }
	 }
	 
    // Reset sheet name
    sheet.setName(new Date().getTime());
    
  } else {
    return Excel.run({ delayForCellEdit: true }, function (context) {
	  var sheet = context.workbook.worksheets.getActiveWorksheet();
	  var now = new Date().getTime();
	  sheet.name = now.toString();
	  sheet.getRange().clear();
      return context.sync();
    })
  }
}


// handles showing popup messages to user
// @param: message - strng of the raw message to show user
// @param: messageTitle - strng of the message title to use
function popupShow(message, messageTitle) {
	if (!message) return;

	if (IS_GOOGLE) {
		var htmlMessage = HtmlService.createHtmlOutput('<p ' + INLINE_STYLING + '>' + message + '</p>').setWidth(200).setHeight(75);
		SpreadsheetApp.getUi().showModalDialog(htmlMessage, messageTitle || "");		 
	} else {
		showPanel("confirm");
		document.getElementById("message-confirm").innerHTML = (messageTitle ? "<b>" + messageTitle + ":</b> " : "") + message;
		document.getElementById("btn-confirm-cancel").style.visibility = "hidden";
		document.getElementById("btn-confirm-ok").onclick = () => popupHide();
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
	var params = { 'content-type': 'application/json' };

	//call Google fetch function (UrlFetchApp) if using google
	if (IS_GOOGLE) {
		var response = UrlFetchApp.fetch(fullUrl, params);
		//returns parsed JSON
		//unparsed response contains error codes if needed
		return JSON.parse(response);

	//for MS Excel, use axios to return a promise to the taskpane
	} else {
		return axios.get(fullUrl);
	}

}



// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
// @param: currentUser - object with details about the current user
function getProjects(currentUser) {
  var fetcherURL = API_BASE_NO_SLASH + '?';
	return fetcher(currentUser, fetcherURL);
}



// Gets components for selected project.
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getComponents(currentUser, projectId) {
	var fetcherURL = API_BASE + projectId + '/components?active_only=true&include_deleted=false&';
	return fetcher(currentUser, fetcherURL);
}



// Gets custom fields for selected project and artifact
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: artifactName - string name of the current artifact
function getCustoms(currentUser, projectId, artifactName) {
	var fetcherURL = API_BASE + projectId + '/custom-properties/' + artifactName + '?';
	return fetcher(currentUser, fetcherURL);
}



// Gets data for a bespoke specified field (for selected project and artifact)
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: artifactName - string name of the current artifact
function getBespoke(currentUser, projectId, artifactName, field) {
	var fetcherURL = API_BASE + projectId + field.bespoke.url + '?';
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
  var fetcherURL = API_BASE + projectId + '/releases?';
  return fetcher(currentUser, fetcherURL);
}



// Gets users for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getUsers(currentUser, projectId) {
  var fetcherURL = API_BASE + projectId + '/users?';
  return fetcher(currentUser, fetcherURL);
}


function getArtifacts(user, projectId, artifactTypeId, startRow, numberOfRows, artifactId) {
  var fullURL = API_BASE + projectId;
  var response = null;
  
    switch (artifactTypeId) {
    case ART_ENUMS.requirements:
      fullURL += "/requirements?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.testCases:
      fullURL += "/test-cases?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.testSteps:
      if (artifactId) {
        fullURL += "/test-cases/" + artifactId + "/test-steps?";
        response = fetcher(user, fullURL);
      }
      break;
    case ART_ENUMS.incidents:
      fullURL += "/incidents/search?start_row=" + startRow + "&number_rows=" + numberOfRows + "&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.releases:
      fullURL += "/releases/search?start_row=" + startRow + "&number_rows=" + numberOfRows + "&";
      var rawResponse = poster("", user, fullURL);
      response = IS_GOOGLE ? JSON.parse(rawResponse) : rawResponse; // this particular return needs to be parsed here
      break;
    case ART_ENUMS.tasks:
      fullURL += "/tasks?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=TaskId&sort_direction=ASC&";
      var rawResponse = poster("", user, fullURL);
      response = IS_GOOGLE ? JSON.parse(rawResponse) : rawResponse; // this particular return needs to be parsed here
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
		//for MS Excel, use axios to return a promise to the taskpane
		// by default axios has a content-type of json, but you get an error when you don't specifically include it in the header
		return axios.post(fullUrl, body, {
			headers: { 'Content-Type': 'application/json' }
		});
	}
}



// effectively a switch to manage which artifact we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactId - int of the current artifact
// @param: parentId - optional int of the relevant parent to attach the artifact too
function postArtifactToSpira(entry, user, projectId, artifactTypeId, parentId) {

	//stringify
	var JSON_body = JSON.stringify(entry),
		response = "",
		postUrl = "";

	//send JSON object of new item to artifact specific export function
	switch (artifactTypeId) {

		// REQUIREMENTS
		case ART_ENUMS.requirements:
			// url to post initial RQ to ensure it is fully outdented
			if (entry.indentPosition === 0 ) { 
				postUrl = API_BASE + projectId + '/requirements/indent/' + INITIAL_HIERARCHY_OUTDENT + '?';
			// if no parentId then post as a regular RQ 
			} else if (parentId === -1) {
				postUrl = API_BASE + projectId + '/requirements?';
			// we should have a parent Id set so add this RQ as its child
			} else {
				postUrl = API_BASE + projectId + '/requirements/parent/' + parentId + '?';
			}
			break;

		// TEST CASES
		case ART_ENUMS.testCases:
			postUrl = API_BASE + projectId + '/test-cases?';
			break;

		// INCIDENTS
		case ART_ENUMS.incidents:
			postUrl = API_BASE + projectId + '/incidents?';
			break;
		
		// RELEASES
		case ART_ENUMS.releases:
			// if no parentId then post as a regular release
			if (parentId === -1) {
				postUrl = API_BASE + projectId + '/releases?';
			// we should have a parent Id set so add this RQ as its child
			} else {
				postUrl = API_BASE + projectId + '/releases/' + parentId + '?';
			}
			break;

		// TASKS
		case ART_ENUMS.tasks:
			postUrl = API_BASE + projectId + '/tasks?';
			break;

		// TEST STEPS
		case ART_ENUMS.testSteps:
			postUrl = parentId !== -1 ? API_BASE + projectId + '/test-cases/' + parentId + '/test-steps?' : null;
			// only post the test step if we have a parent id
			break;
	}

	return postUrl ? poster(JSON_body, user, postUrl) : null;
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
function error(type) {
	if (type == 'impExp') {
		okWarn('There was an input error. Please check that your entries are correct.');
	} else if (type == 'unknown') {
		okWarn('Unkown error. Please try again later or contact your system administrator');
	} else {
		okWarn('Network error. Please check your username, url, and password. If correct make sure you have the correct permissions.');
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
function operationComplete(messageEnum) {
	if (IS_GOOGLE) {
		var message = STATUS_MESSAGE_GOOGLE[messageEnum] || STATUS_MESSAGE_GOOGLE['1'];
		okWarn(message);
	} else {
		var message = STATUS_MESSAGE_EXCEL[messageEnum] || STATUS_MESSAGE_EXCEL['1'];
		popupShow(message, "");
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
// @param: fieldType - list of fieldType enums from client params object
function templateLoader(model, fieldType) {
	// clear spreadsheet depending on user input, and unhide everything
	clearAll();
  
  var fields = model.fields;
  var sheet;
  var newSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;

  // select active sheet
  if (IS_GOOGLE) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // set sheet (tab) name to model name
    sheet.setName(newSheetName);
    sheetSetForTemplate(sheet, model, fieldType, null);

  } else {
    return Excel.run(function (context) {
	  sheet = context.workbook.worksheets.getActiveWorksheet();
	  sheet.name = newSheetName;
      return context.sync()
        .then(function() {
          sheetSetForTemplate(sheet, model, fieldType, context); 
        }).catch(/*fail quietly*/);
    })
  } 
}

// wrapper function to set the header row, validation rules, and any extra formatting
function sheetSetForTemplate(sheet, model, fieldType, context) {
  // heading row - sets names and formatting
  headerSetter(sheet, model.fields, model.colors, context);
  // set validation rules on the columns
  contentValidationSetter(sheet, model, fieldType, context);
  // set any extra formatting options
  contentFormattingSetter(sheet, model, context);   
}



// Sets headings for fields
// creates an array of the field names so that changes can be batched to the relevant range in one go for performance reasons
// @param: sheet - the sheet object
// @param: fields - full field data
// @param: colors - global colors used for formatting
function headerSetter (sheet, fields, colors, context) {

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
// @param: fieldType - enums for field types
function contentValidationSetter (sheet, model, fieldType, context) {
  // we can't easily get the max rows for excel so use the number of rows it always seems to have
  var nonHeaderRows =  IS_GOOGLE ? sheet.getMaxRows() - 1 : 1048576 - 1;

	for (var index = 0; index < model.fields.length; index++) {
		var columnNumber = index + 1,
			list = [];

		switch (model.fields[index].type) {

			// ID fields: restricted to numbers and protected
			case fieldType.id:
			case fieldType.subId:
				setNumberValidation(sheet, columnNumber, nonHeaderRows, false);
				protectColumn(
					sheet,
					columnNumber,
					nonHeaderRows,
					model.colors.bgReadOnly,
					"ID field",
					false
					);
				break;

			// INT and NUM fields are both treated by Sheets as numbers
			case fieldType.int:
			case fieldType.num:
				setNumberValidation(sheet, columnNumber, nonHeaderRows, false);
				break;

			// BOOL as Sheets has no bool validation, a yes/no dropdown is used
			case fieldType.bool:
				// 'True' and 'False' don't work as dropdown choices
				list.push("Yes", "No");
				setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
				break;
			
			// DATE fields get date validation
			case fieldType.date:
				setDateValidation(sheet, columnNumber, nonHeaderRows, false);
				break;
			
			// DROPDOWNS and MULTIDROPDOWNS are both treated as simple dropdowns (Sheets does not have multi selects)
			case fieldType.drop:
			case fieldType.multi:
				var fieldList = model.fields[index].values;
				for (var i = 0; i < fieldList.length; i++) {
					list.push(fieldList[i].name);
				}
				setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
				break;

			// USER fields are dropdowns with the values coming from a project wide set list
			case fieldType.user:
				for (var j = 0; j < model.projectUsers.length; j++) {
					list.push(model.projectUsers[j].name);
				}
				setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
				break;

			// COMPONENT fields are dropdowns with the values coming from a project wide set list
			case fieldType.component:
				for (var k = 0; k < model.projectComponents.length; k++) {
					list.push(model.projectComponents[k].name);
				}
				setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
				break;

			// RELEASE fields are dropdowns with the values coming from a project wide set list
			case fieldType.release:
				for (var l = 0; l < model.projectReleases.length; l++) {
					list.push(model.projectReleases[l].name);
				}
				setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
				break;

			// All other types
			default:
				//do nothing
				break;
		}
  }
}



// create dropdown validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: list - array of values to show in a dropdown and use for validation
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setDropdownValidation (sheet, columnNumber, rowLength, list, allowInvalid) {
  if (IS_GOOGLE) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);
  
    // create the validation rule
    // requireValueInList - params are the array to use, and whether to create a dropdown list
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(allowInvalid)
      .build();
    range.setDataValidation(rule);
  
  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();
    let approvedListRule = {
        list: {
            inCellDropDown: true,
            source: list.join()
        }
    };
    range.dataValidation.rule = approvedListRule;
  }
}



// create date validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setDateValidation (sheet, columnNumber, rowLength, allowInvalid) {
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

  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    let greaterThan1970Rule = {
      date: {
          formula1: "2000-01-01",
          operator: Excel.DataValidationOperator.greaterThan
      }
    };
    range.dataValidation.rule = greaterThan1970Rule;

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
  }
}



// create number validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setNumberValidation (sheet, columnNumber, rowLength, allowInvalid) {
  // create range
  if (IS_GOOGLE) {
    var range = sheet.getRange(2, columnNumber, rowLength);
  
    // create the validation rule
    //must be a valid number greater than -1 (also excludes 1.1.0 style numbers)
    var rule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(-1)
      .setAllowInvalid(allowInvalid)
      .setHelpText('Must be a positive number')
      .build();
    range.setDataValidation(rule);
  
  } else {
    var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
    range.dataValidation.clear();

    let greaterThanZeroRule = {
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


// format columns based on a potential rang of factors - eg hide unsupported columns
// @param: sheet - the sheet object
// @param: model - full model data set
function contentFormattingSetter (sheet, model) {
	for (var i = 0; i < model.fields.length; i++) {
		var columnNumber = i + 1;

		// change bg color of unsupported fields
		if (model.fields[i].unsupported) {
			protectColumn(
			  sheet,
			  columnNumber,
			  (sheet.getMaxRows() - 1),
			  model.colors.bgReadOnly,
			  model.fields[i].name + "unsupported",
			  true
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
// @param: hide - optional bool to hide column completely
function protectColumn (sheet, columnNumber, rowLength, bgColor, name, hide) {
  // only for google as cannot protect individual cells easily in Excel
  if (IS_GOOGLE) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);
    range.setBackground(bgColor)
      .protect()
      .setDescription(name)
      .setWarningOnly(true);
  
    if(hide) {
    sheet.hideColumns(columnNumber);
    }
  } else {
	var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
	
	// set the background color
	range.set({ format: { fill: { color: bgColor } } });
	
	// now we can add data validation
	// the easiest hack way to not allow any entry into the cell is to make sure its text length can only be zero
	range.dataValidation.clear();
    let textLengthZero = {
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









/*
 * ================
 * SENDING TO SPIRA
 * ================
 *
 * The main function takes the entire data model and the artifact type
 * and calls the child function to set various object values before
 * sending the finished objects to SpiraTeam
 *
 */

// function that manages exporting data from the sheet - creating an array of objects based on entered data, then sending to Spira
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldType - list of fieldType enums from client params object
function sendToSpira(model, fieldType) {
	// 0. SETUP FUNCTION LEVEL VARS
	var fields = model.fields,
		artifact = model.currentArtifact,
		artifactIsHierarchical = artifact.hierarchical,
		requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id
	
	// 1. get the active spreadsheet and first sheet
	if (IS_GOOGLE) {
		var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
			sheet = spreadSheet.getActiveSheet(),
			lastRow = sheet.getLastRow() - 1 || 10, // hack to make sure we pass in some rows to the sheetRange, otherwise it causes an error
			sheetRange = sheet.getRange(2,1, lastRow, fields.length),
			sheetData = sheetRange.getValues(),
			entriesForExport = createExportEntries(sheetData, model, fieldType, fields, artifact, artifactIsHierarchical);
		if (sheet.getName() == requiredSheetName) {
			return sendExportEntriesGoogle(entriesForExport, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, null);
		} else {
			var log = {
				status: STATUS_ENUM.wrongSheet 
			};
			return log;
		}
	} else {
		return Excel.run({ delayForCellEdit: true }, function (context) {
			var sheet = context.workbook.worksheets.getActiveWorksheet(),
				sheetRange = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS, fields.length);
			sheet.load("name");
			sheetRange.load("values");
			return context.sync()
		.then(() => {
			if (sheet.name == requiredSheetName) {
				var sheetData = sheetRange.values,
					entriesForExport = createExportEntries(sheetData, model, fieldType, fields, artifact, artifactIsHierarchical);
				return sendExportEntriesExcel(entriesForExport, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, context);
			} else {
				var log = {
					status: STATUS_ENUM.wrongSheet 
				};
				return log;
			}
		})
		.catch();
		}).catch();
	}
}



// 2. CREATE ARRAY OF ENTRIES
// loop to create artifact objects from each row taken from the spreadsheet
// vars needed: sheetData, artifact, fields, model, fieldType, artifactIsHierarchical,
function createExportEntries(sheetData, model, fieldType, fields, artifact, artifactIsHierarchical) {
	var lastIndentPosition = null,
		entriesForExport = [];

	for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
	
		// stop at the first row that is fully blank
		if (sheetData[rowToPrep].join("") === "") {
			break;
		} else {
			// check for required fields (for normal artifacts and those with sub types - eg test cases and steps)
			var rowChecks = {
					hasSubType: artifact.hasSubType,
					totalFieldsRequired: countRequiredFieldsByType(fields, false),
					totalSubTypeFieldsRequired: artifact.hasSubType ? countRequiredFieldsByType(fields, true) : 0,
					countRequiredFields: rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, false),
					countSubTypeRequiredFields: artifact.hasSubType ? rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, true) : 0,
					subTypeIsBlocked: !artifact.hasSubType ? true : rowBlocksSubType(sheetData[rowToPrep], fields)
				},
	
				// create entry used to populate all relevant data for this row
				entry = {};
	
			// first check for errors
			var hasProblems = rowHasProblems(rowChecks);
			if (hasProblems) {
				entry.validationMessage = hasProblems;
	
			// if error free determine what field filtering is required - needed to choose type/subtype fields if subtype is present
			} else {
				var fieldsToFilter = relevantFields(rowChecks);
				entry = createEntryFromRow( sheetData[rowToPrep], model, fieldType, artifactIsHierarchical, lastIndentPosition, fieldsToFilter );
	
				// FOR SUBTYPE ENTRIES add flag on entry if it is a subtype
				if (fieldsToFilter === FIELD_MANAGEMENT_ENUMS.subType) {
					entry.isSubType = true;
				}
				// FOR HIERARCHICAL ARTIFACTS update the last indent position before going to the next entry to make sure relative indent is set correctly
				if (artifactIsHierarchical) {
					lastIndentPosition = entry.indentPosition;
				}
			}
			entriesForExport.push(entry);
		}
	}
	return entriesForExport;
}


// 3. FOR GOOGLE ONLY: GET READY TO SEND DATA TO SPIRA + 4. ACTUALLY SEND THE DATA
// check we have some entries and with no errors
// Create and show a message to tell the user what is going on
function sendExportEntriesGoogle(entriesForExport, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, context) {
	if (!entriesForExport.length) {
		popupShow('There are no entries to send to Spira', 'Check Sheet')
		return "nothing to send";
	} else {
		popupShow('Preparing to send...', 'Progress');

		// create required variables for managing responses for sending data to spirateam
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
			var sentToSpira = manageSendingToSpira(entriesForExport[i], model.user, model.currentProject.id, artifact, fields, fieldType, log.parentId );

			// update the parent ID for a subtypes based on the successful API call
			if (artifact.hasSubType) {
				log.parentId = sentToSpira.parentId;
			}

			log = processSendToSpiraResponse(i, sentToSpira, entriesForExport, artifact, response, log);
			if (sentToSpira.error && artifact.hierarchical) {
				// break out of the recursive loop
				log.doNotContinue = true;
			}
		}

		// 4. SEND DATA TO SPIRA AND MANAGE RESPONSES
		// KICK OFF THE FOR LOOP (IE THE FUNCTION ABOVE) HERE
		// We use a function rather than a loop so that we can more readily use promises to chain things together and make the calls happen synchronously
		// we need the calls to be synchronous because we need to do the status and ID of the preceding entry for hierarchical artifacts
		for (var i = 0; i < entriesForExport.length; i++) {
			if (!log.doNotContinue) {
				log = checkSingleEntryForErrors(i, log, entriesForExport, artifact);
				if (log.entries.length && log.entries[i] && log.entries[i].error) {
				} else {
					sendSingleEntry(i);	
				}
			}
		}

		// review all activity and set final status
		log.status = log.errorCount ? (log.errorCount == log.entriesLength ? STATUS_ENUM.allError : STATUS_ENUM.someError) : STATUS_ENUM.allSuccess;

		// call the final function here - so we know that it is only called after the recursive function above (ie all posting) has ended
		return updateSheetWithExportResults(log, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, context);
 	}
}


// 3. FOR EXCEL ONLY: GET READY TO SEND DATA TO SPIRA + 4. ACTUALLY SEND THE DATA
// DIFFERENT TO GOOGLE: this uses async await for its function and subfunction
// check we have some entries and with no errors
// Create and show a message to tell the user what is going on
async function sendExportEntriesExcel(entriesForExport, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, context) {
	if (!entriesForExport.length) {
		popupShow('There are no entries to send to Spira', 'Check Sheet')
		return "nothing to send";
	} else {
		popupShow('Starting to send...', 'Progress');

		// create required variables for managing responses for sending data to spirateam
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
			
			await manageSendingToSpira(entriesForExport[i], model.user, model.currentProject.id, artifact, fields, fieldType, log.parentId )
				.then(function(response) {
					// update the parent ID for a subtypes based on the successful API call
					if (artifact.hasSubType) {
						log.parentId = sentToSpira.parentId;
					}
					console.log('manageSendingToSpira > then for ', i, response);
					log = processSendToSpiraResponse(i, response, entriesForExport, artifact, response, log);

					if (response.error && artifact.hierarchical) {
						// break out of the recursive loop
						log.doNotContinue = true;
					}
				})
		}



		// 4. SEND DATA TO SPIRA AND MANAGE RESPONSES
		// KICK OFF THE FOR LOOP (IE THE FUNCTION ABOVE) HERE
		// We use a function rather than a loop so that we can more readily use promises to chain things together and make the calls happen synchronously
		// we need the calls to be synchronous because we need to do the status and ID of the preceding entry for hierarchical artifacts
		for (var i = 0; i < entriesForExport.length; i++) {
			if (!log.doNotContinue) {
				log = checkSingleEntryForErrors(i, log, entriesForExport, artifact);
				if (log.entries.length && log.entries[i] && log.entries[i].error) {
				} else {
					await sendSingleEntry(i);	
				}
			}
		}

		// review all activity and set final status
		log.status = log.errorCount ? (log.errorCount == log.entriesLength ? STATUS_ENUM.allError : STATUS_ENUM.someError) : STATUS_ENUM.allSuccess;

		// call the final function here - so we know that it is only called after the recursive function above (ie all posting) has ended
		return updateSheetWithExportResults(log, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, context);
 	}
}


// 5. SET MESSAGES AND FORMATTING ON SHEET
function updateSheetWithExportResults(log, sheetData, sheet, sheetRange, model, fieldType, fields, artifact, context) {
	var bgColors = [],
		notes = [],
		values = [];
	// first handle cell formatting
	for (var row = 0; row < log.entries.length; row++) {
		var rowBgColors = [],
			rowNotes = [],
			rowValues = [];
		for (var col = 0; col < fields.length; col++) {
			var bgColor,
				note = null,
				value = sheetData[row][col];
			
				// first handle when we are dealing with data that has been sent to Spira
			var isSubType = (log.entries[row].details &&  log.entries[row].details.entry && log.entries[row].details.entry.isSubType) ? log.entries[row].details.entry.isSubType : false;
			
			bgColor = setFeedbackBgColor(sheetData[row][col], log.entries[row].error, fields[col], fieldType, artifact, model.colors);
			note = setFeedbackNote(sheetData[row][col], log.entries[row].error, fields[col], fieldType, log.entries[row].message);
			value = setFeedbackValue(sheetData[row][col], log.entries[row].error, fields[col], fieldType, log.entries[row].newId || "", isSubType);

			if (IS_GOOGLE) {
				rowBgColors.push(bgColor);
				rowNotes.push(note);
				rowValues.push(value);
			} else {
				var cellRange = sheet.getCell(row + 1, col);
				if (note) rowNotes.push(note);
				if (bgColor) {
					cellRange.set({ format: {fill: { color: bgColor }} });
				} else {
					cellRange.clear();
				}
				cellRange.values = [[value]];

			}
		}
		if (IS_GOOGLE) {
			bgColors.push(rowBgColors);
			notes.push(rowNotes);
			values.push(rowValues);
		
		// for Excel we can't pass in arrays of data for values, but we still take action here for notes - because Excel API does not allow the addition of comments
		} else {
			var endRowCell = sheet.getCell(row + 1, fields.length);
			if (rowNotes.length) {
				endRowCell.set({ format: {fill: { color: model.colors.warning }} });
				endRowCell.values = [[rowNotes.join()]];

				var endRowHeader = sheet.getCell(0, fields.length);
				endRowHeader.set({ format: {
					fill: { color: model.colors.bgHeader },
					font: { color: model.colors.header }
				}});
				endRowHeader.values = [["Error Messages"]];
			} else {
				endRowCell.clear();
			}
		}
	}	

	if (IS_GOOGLE) {
		sheetRange.setBackgrounds(bgColors).setNotes(notes).setValues(values);
		return log;
	} else {
		console.log('excel thinks the whole thing is finished')
		return context.sync().then(() => log);
	}
}



function checkSingleEntryForErrors(i, log, entriesForExport, artifact, parentId) {
	var response = {};

	// skip if there was an error validating the sheet row
	if (entriesForExport[i].validationMessage) {
		response.error = true;
		response.message = entriesForExport[i].validationMessage;
		log.errorCount++;

		// stop if the artifact is hierarchical because we don't know what side effects there could be to any further items.
		if (artifact.hierarchical) {
			log.doNotContinue = true;
			response.message += " - no further entries were sent to avoid creating an incorrect hierarchy";
			// make sure to push the response so that the client can process error message
			log.entries.push(response);
			// we do not call this function again with i++ so that we effectively break out of the loop
			// review all activity and set final status
			log.status = log.errorCount ? (log.errorCount == log.entriesLength ? STATUS.allError : STATUS_ENUM.someError) : STATUS_ENUM.allSuccess;
		} else {
			log.entries.push(response);
		}
	// skip if a sub type row does not have a parent to hook to
	} else if (entriesForExport[i].isSubType && !log.parentId) {
		response.error = true;
		response.message = "can't add a child type when there is no corresponding parent type";
		log.errorCount++;
		log.entries.push(response);
	}
	return log;
}



// function that reviews a specific cell against it's field and errors for providing UI feedback on errors
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldType - enum information about field types
// @param: artifact - the currently selected artifact
// @param: colors - object of colors to use based on different conditions
function setFeedbackBgColor (cell, error, field, fieldType, artifact, colors) {
	if (error) {
		// if we have a validation error, we can highlight the relevant cells if the art has no sub type
		if (!artifact.hasSubType) {
			if (field.required && cell === "") {
				return colors.warning;
			} else {
				// keep original formatting
				if (field.type == fieldType.subId || field.type == fieldType.id || field.unsupported) {
					return colors.bgReadOnly;
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
		if (field.type == fieldType.subId || field.type == fieldType.id || field.unsupported) {
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
// @param: fieldType - enum information about field types
// @param: message - relevant error message from the entry for this row
function setFeedbackNote (cell, error, field, fieldType, message) {
	// handle entries with errors - add error notes into ID field
	if (error && field.type == fieldType.id) {
		return message;
	} else {
		return null;
	}
}



// function that updates id fields with new values, otherwise returns existing value
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldType - enum information about field types
// @param: newId - int that is the newly created Id for this row
// @param: isSubType - bool if row is subtype or not - false on error as there will be no id to add anyway
function setFeedbackValue (cell, error, field, fieldType, newId, isSubType) {
	// when there is an error we don't change any of the cell data
	if (error) {
		return cell;

		// handle successful entries - ie add ids into right place
	} else {
		var newIdToEnter =	newId || "";
		if (!isSubType && field.type == fieldType.id) {
			return newIdToEnter;
		} else if (isSubType && field.type == fieldType.subId) {
			return newIdToEnter;
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
// @param: fieldType - object of all field types with enums
function manageSendingToSpira (entry, user, projectId, artifact, fields, fieldType, parentId) {
	var data,
		// set output parent id here so we know this function will always return a value for this
		output = {
			parentId: parentId,
			entry: entry,
			artifact: {
				artifactId: artifactTypeIdToSend,
				artifactObject: artifact
			}
		},
		// make sure correct artifact ID is sent to handler (ie type vs subtype)
		artifactTypeIdToSend = entry.isSubType ? artifact.subTypeId : artifact.id;

	
	// send object to relevant artifact post service
	if (IS_GOOGLE) {
		data = postArtifactToSpira(entry, user, projectId, artifactTypeIdToSend, parentId);

		// save data for logging to client
		output.httpCode = (data && data.getResponseCode() ) ? data.getResponseCode() : "notSent";

		// parse the data if we have a success
		if (output.httpCode == 200) {
			output.fromSpira = JSON.parse(data.getContentText());

			// get the id/subType id of the newly created artifact
			var artifactIdField = getIdFieldName(fields, fieldType, entry.isSubType);
			output.newId = output.fromSpira[artifactIdField];

			// update the output parent ID to the new id only if the artifact has a subtype and this entry is NOT a subtype
			if (artifact.hasSubType && !entry.isSubType) {
				output.parentId = output.newId;
			}

		} else {
			//we have an error - so set the flag and the message
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
		return output;

	} else {
		return postArtifactToSpira(entry, user, projectId, artifactTypeIdToSend, parentId)
			.then((response) => { 
				output.fromSpira = response.data;

				// get the id/subType id of the newly created artifact
				var artifactIdField = getIdFieldName(fields, fieldType, entry.isSubType);
				output.newId = output.fromSpira[artifactIdField];

				// update the output parent ID to the new id only if the artifact has a subtype and this entry is NOT a subtype
				if (artifact.hasSubType && !entry.isSubType) {
					output.parentId = output.newId;
				} 
				return output;
			})
			.catch((error) => { 
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



// returns the correct parentId for the relevant indent position by looping back through the list of entries
// returns -1 if no match found
// @param: indent - int of the indent position to retrieve the parent for
// @param: previousEntries - object containing all successfully sent entries - with, if a hierarchical artifact, a hierarchy info object
function getHierarchicalParentId (indent, previousEntries) {
	// if there is no indent/ set to initial indent we return out immediately 
	if (indent === 0 || !previousEntries.length) {
		return -1;
	}
	for (var i = previousEntries.length - 1; i >=0; i--) {
		// when the indent is greater - means we are indenting, so take the last array item and return out
		// check for presence of correct objects in item - should exist, as otherwise error should be thrown
		if (previousEntries[i].details && previousEntries[i].details.hierarchyInfo && previousEntries[i].details.hierarchyInfo.indent < indent) {
			return previousEntries[i].details.hierarchyInfo.id;
			break;
		}
	}
	return -1;
}


// returns an int of the total number of required fields for the passed in artifact
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
// @param: forSubType - bool to determine whether to check for sub type required fields (true), or not - defaults to false
function countRequiredFieldsByType (fields, forSubType) {
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
function rowCountRequiredFieldsByType (row, fields, forSubType) {
	var count = 0;
	for (var i = 0; i < row.length; i++) {
		if (forSubType != "undefined" && forSubType) {
			if (fields[i].requiredForSubType && row[i]) {
				count++;
			}
		} else if (fields[i].required && row[i]) {
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
function rowBlocksSubType (row, fields) {
	var result = false;
	for (var column = 0; column < row.length; column++) {
		if (fields[column].forbidOnSubType && row[column]) {
			result = true;
		}
	}
	return result;
}



// checks to see if the row is valid - ie required fields present and correct as expected
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: rowChecks - object with different properties for different checks required
function rowHasProblems (rowChecks) {
	var problems = null;
	if (!rowChecks.hasSubType && rowChecks.countRequiredFields < rowChecks.totalFieldsRequired) {
		problems = "Fill in all required fields";
	} else if (rowChecks.hasSubType) {
		if (rowChecks.countSubTypeRequiredFields < rowChecks.totalSubTypeFieldsRequired && !rowChecks.countRequiredFields) {
			problems = "Fill in all required fields";
		} else if (rowChecks.countRequiredFields < rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
			problems = "Fill in all required fields";
		} else if (rowChecks.countSubTypeRequiredFields == rowChecks.totalSubTypeFieldsRequired && (rowChecks.countRequiredFields == rowChecks.totalFieldsRequired || rowChecks.subTypeIsBlocked) ){
			problems = "It is unclear what artifact this is intended to be";
		}
	}
	return problems;
}



// based on field type and conditions, determines what fields are required for a given row
// e.g. all fields is default and standard, if a subtype is present (eg test step) - should it send only the main type or the sub type fields
// returns a int representing the relevant enum value
// @ param: rowChecks - object with different properties for different checks required
function relevantFields (rowChecks) {
	var fields = FIELD_MANAGEMENT_ENUMS.all;
	if (rowChecks.hasSubType) {
		if (rowChecks.countRequiredFieldsFilled == rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
			fields = FIELD_MANAGEMENT_ENUMS.standard;
		} else if (rowChecks.countSubTypeRequiredFields == rowChecks.totalSubTypeFieldsRequired && !(rowChecks.countRequiredFields == rowChecks.totalFieldsRequired || rowChecks.subTypeIsBlocked) ) {
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
// @param: fieldType - object of all field types with enums
// @param: artifactIsHierarchical - bool to tell function if this artifact has hierarchy (eg RQ and RL)
// @param: lastIndentPosition - int used for calculating relative indents for hierarchical artifacts
// @param: fieldsToFilter - enum used for selecting fields to not add to object - defaults to using all if omitted
function createEntryFromRow (row, model, fieldType, artifactIsHierarchical, lastIndentPosition, fieldsToFilter) {
	//create empty 'entry' object - include custom properties array here to avoid it being undefined later if needed
	var entry = {
			"CustomProperties": []
		},
		fields = model.fields;

	//we need to turn an array of values in the row into a validated object
	for (var index = 0; index < row.length; index++) {

		// first ignore entry that does not match the requirement specified in the fieldsToFilter
		if (fieldsToFilter == FIELD_MANAGEMENT_ENUMS.standard && fields[index].isSubTypeField ) {
			// skip the field
		} else if (fieldsToFilter == FIELD_MANAGEMENT_ENUMS.subType && !(fields[index].isSubTypeField || fields[index].isTypeAndSubTypeField) ) {
			// skip the field

		// in all other cases add the field
		} else {
			var value = null,
				customType = "",
				idFromName = 0;

			// double check data validation, convert dropdowns to required int values
			// sets both the value, and custom types - so that custom fields are handled correctly
			switch (fields[index].type) {

				// ID fields: restricted to numbers and blank on push, otherwise put
				case fieldType.id:
				case fieldType.subId:

					customType = "IntegerValue";
					break;

				// INT and NUM fields are both treated by Sheets as numbers
				case fieldType.int:
				case fieldType.num:
					// only set the value if a number has been returned
					if (!isNaN(row[index])) {
						value = row[index];
						customType = "IntegerValue";
					}
					break;

				// BOOL as Sheets has no bool validation, a yes/no dropdown is used
				case fieldType.bool:
					// 'True' and 'False' don't work as dropdown choices, so have to convert back
					if (row[index] == "Yes") {
						value = true;
						customType = "BooleanValue";
					} else if (row[index] == "No") {
						value = false;
						customType = "BooleanValue";
					}
					break;

				// DATES - parse the data and add prefix/suffix for WCF
				case fieldType.date:
					if (row[index]) {
						value = "\/Date(" + Date.parse(row[index]) + ")\/";
						customType = "DateTimeValue";
					}
					break;
							
				// ARRAY fields are for multiselect lists - currently not supported so just push value into an array to make sure server handles it correctly
				case fieldType.arr:
					if (row[index]) {
						value = [row[index]];
						customType = ""; // array fields not used for custom properties here
					}
					break;
				
				// DROPDOWNS - get id from relevant name, if one is present
				case fieldType.drop:
					idFromName = getIdFromName(row[index], fields[index].values);
					if (idFromName) {
						value = idFromName;
						customType = "IntegerValue";
					}
					break;

				// MULTIDROPDOWNS - get id from relevant name, if one is present, set customtype to list value
				case fieldType.multi:
					idFromName = getIdFromName(row[index], fields[index].values);
					if (idFromName) {
						value = idFromName;
						customType = "IntegerListValue";
					}
					break;

				// USER fields - get id from relevant name, if one is present
				case fieldType.user:
					idFromName = getIdFromName(row[index], model.projectUsers);
					if (idFromName) {
						value = idFromName;
						customType = "IntegerValue";
					}
					break;

				// COMPONENT fields - get id from relevant name, if one is present
				case fieldType.component:
					idFromName = getIdFromName(row[index], model.projectComponents);
					if (idFromName) {
						value = idFromName;
						customType = "IntegerValue";
					}
					break;

				// RELEASE fields - get id from relevant name, if one is present
				case fieldType.release:
					idFromName = getIdFromName(row[index], model.projectReleases);
					if (idFromName) {
						value = idFromName;
						customType = "IntegerValue";
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
				var indentCount = countIndentCharacters(value, model.indentCharacter);
				var indentPosition = setRelativePosition(indentCount, lastIndentPosition);

				// make sure to slice off the indent characters from the front
				// TODO should also trim white space at start
				value = value.slice(indentCount, value.length);

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
				entry[fields[index].field] = value;
			}
		}

	}

	return entry;
}



// find the corresponding ID for a string value - eg from a dropdown
// dropdowns can only contain one item per row so we have to now get the IDs for sending to Spira
// @param: string - the string of the name value specified
// @param: list - the array of items with keys for id and name values
function getIdFromName (string, list) {
	for (var i = 0; i < list.length; i++) {
		if (list[i].name == string) {
			return list[i].id;
		}
	}
	// return 0 if there's no match
	return 0;
}


// finds and returns the field name for the specific artifiact's ID field
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldType - object of all field types with enums
// @param: getSubType - optioanl bool to specify to return the subtype Id field, not the normal field (where two exist)
function getIdFieldName (fields, fieldType, getSubType) {
	for (var i = 0; i < fields.length; i++) {
		var fieldToLookup = getSubType ? "subId" : "id";
		if (fields[i].type == fieldType[fieldToLookup]) {
			return fields[i].field;
		}
	}
	return null;
}


// returns the count of the number of indent characters and returns the value
// @param: field - a single field string - one already designated as containing hierarchy info
// @param: indentCharacter - the character used to denote an indent - e.g. ">"
function countIndentCharacters (field, indentCharacter) {
	var indentCount = 0;
	//check for field value and indent character
	if (field && field[0] === indentCharacter) {
		//increment indent counter while there are '>'s present
		while (field[0] === indentCharacter) {
			//get entry length for slice
			var len = field.length;
			//slice the first character off of the entry
			field = field.slice(1, len);
			indentCount++;
		}
	}
	return indentCount;
}



// returns the correct relative indent position - based on the previous relative indent and other logic (int neg, pos, or zero)
// Currently the API does not support a call to place an artifact at a certain location.
// @param: indentCount - int of the number of indent characters set by user
// @param: lastIndentPosition - int sum of the actual indent positions used for the preceding entries
function setRelativePosition (indentCount, lastIndentPosition) {
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
function processSendToSpiraResponse(i, sentToSpira, entriesForExport, artifact, response, log) {
	
	response.details = sentToSpira;

	// handle success and error cases
	if (sentToSpira.error) {
		log.errorCount++;
		response.error = true;
		response.message = sentToSpira.errorMessage;
	
		//Sets error HTML modals
		if (artifact.hierarchical) {
			// if there is an error on any hierarchical artifact row, break out of the loop to prevent entries being attached to wrong parent
			popupShow('Error sending ' + (i + 1) + ' of ' + (entriesForExport.length) + ' - sending stopped to avoid indenting entries incorrectly', 'Progress')
			log.entries.push(response);
		} else { 
			popupShow('Error sending ' + (i + 1) + ' of ' + (entriesForExport.length), 'Progress');
		}

	} else {
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
		popupShow('Sent ' + (i + 1) + ' of ' + (entriesForExport.length) + '...', 'Progress');
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
// @param: model: full model object from client
// @param: enum of fieldTypes used
function getFromSpiraGoogle(model, fieldType) {
  //var artifactCount = getArtifactCount(model.user, model.currentProject.id, model.currentArtifact.id);

  // 1. get from spira
  // note we don't do this by getting the count of each artifact first, because of a bug in getting the release count
  var currentPage = 0;
  var artifacts = [];
  var getNextPage = true;
  
  while (getNextPage && currentPage < 100) {
    var startRow = (currentPage * GET_PAGINATION_SIZE) + 1;
    var pageOfArtifacts = getArtifacts(
      model.user, 
      model.currentProject.id, 
      model.currentArtifact.id, 
      startRow, 
      GET_PAGINATION_SIZE,
      null
    );
    // if we got a non empty array back then we have artifacts to process
    if (pageOfArtifacts.length) {
      artifacts = artifacts.concat(pageOfArtifacts);
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
    artifacts.sort(function(a, b) {
      return a.indentLevel < b.indentLevel ? -1 : 1;
    });
  }
  
  
  // 4. if artifact has subtype that needs to be retrieved separately, do so
  if (model.currentArtifact.hasSubType) {
    // find the id field
    var idFieldNameArray = model.fields.filter(function(field) {
      return field.type === fieldType.id;
    });
    // if we have an id field, then we can find the id number for each artifact in the array
    if (idFieldNameArray && idFieldNameArray[0].field) {
      var idFieldName = idFieldNameArray[0].field;
      var artifactsWithSubTypes = [];
      artifacts.forEach(function(art) {
        artifactsWithSubTypes.push(art);
        var subTypeArtifacts = getArtifacts(
          model.user, 
          model.currentProject.id, 
          model.currentArtifact.subTypeId, 
          null, 
          null,
          art[idFieldName]
        );
        // take action if we got any sub types back - ie if they exist for the specific artifact
        if (subTypeArtifacts && subTypeArtifacts.length) { 
          var subTypeArtifactsWithMeta = subTypeArtifacts.map(function(sub) {
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
    fieldType, 
    model.projectUsers, 
    model.projectComponents, 
    model.projectReleases
  );
  
  // 6. add data to sheet
  if (IS_GOOGLE) {
	  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = spreadSheet.getActiveSheet(),
		range = sheet.getRange(2,1, artifacts.length, model.fields.length);
	  
	  range.setValues(artifactsAsCells);
	  return JSON.parse(JSON.stringify(artifactsAsCells));
  } else {
	return Excel.run({ delayForCellEdit: true }, function (context) {
		var sheet = context.workbook.worksheets.getActiveWorksheet(),
			range = sheet.getRangeByIndexes(1, 0, artifacts.length, model.fields.length);
		range.values = artifactsAsCells;
		return context.sync()
			.then(function() {
				return artifactsAsCells;
			})
	  })
  }
}

// EXCEL SPECIFIC VARIATION OF THIS FUNCTION handles getting paginated artifacts from Spira and saving them as a single array
// @param: model: full model object from client
// @param: enum of fieldTypes used
async function getFromSpiraExcel(model, fieldType) {
	// 1. get from spira
	// note we don't do this by getting the count of each artifact first, because of a bug in getting the release count
	var currentPage = 0;
	var artifacts = [];
	var getNextPage = true;

	async function getArtifactsPage(startRow) {
		await getArtifacts(
			model.user, 
			model.currentProject.id, 
			model.currentArtifact.id, 
			startRow, 
			GET_PAGINATION_SIZE,
			null
		).then(function(response) {
			// if we got a non empty array back then we have artifacts to process
			if (response.data.length) {
				artifacts = artifacts.concat(response.data);
				// if we got less artifacts than the max we asked for, then we reached the end of the list in this request - and should stop
				if (response.data.length < GET_PAGINATION_SIZE) {
				getNextPage = false;
				// if we got the full page size back then there may be more artifacts to get
				} else {
				currentPage++;
				}
			// if we got no artifacts back, stop now
			} else {
				getNextPage = false;
			}
		})
	}
	
	while (getNextPage && currentPage < 100) {
	  var startRow = (currentPage * GET_PAGINATION_SIZE) + 1;
	  await getArtifactsPage(startRow);
	  // if we got a non empty array back then we have artifacts to process
	}
	
	// 2. if there were no artifacts at all break out now
	if (!artifacts.length) return "no artifacts were returned";
	
	// 3. Make sure hierarchical artifacts are ordered correctly
	if (model.currentArtifact.hierarchical) {
	  artifacts.sort(function(a, b) {
		return a.indentLevel < b.indentLevel ? -1 : 1;
	  });
	}
	
	
	// 4. if artifact has subtype that needs to be retrieved separately, do so
	if (model.currentArtifact.hasSubType) {
	  // find the id field
	  var idFieldNameArray = model.fields.filter(function(field) {
		return field.type === fieldType.id;
	  });
	  // if we have an id field, then we can find the id number for each artifact in the array
	  if (idFieldNameArray && idFieldNameArray[0].field) {
		var idFieldName = idFieldNameArray[0].field;
		var artifactsWithSubTypes = [];
		artifacts.forEach(function(art) {
		  artifactsWithSubTypes.push(art);
		  var subTypeArtifacts = getArtifacts(
			model.user, 
			model.currentProject.id, 
			model.currentArtifact.subTypeId, 
			null, 
			null,
			art[idFieldName]
		  );
		  // take action if we got any sub types back - ie if they exist for the specific artifact
		  if (subTypeArtifacts && subTypeArtifacts.length) { 
			var subTypeArtifactsWithMeta = subTypeArtifacts.map(function(sub) {
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
	  fieldType, 
	  model.projectUsers, 
	  model.projectComponents, 
	  model.projectReleases
	);
	
	// 6. add data to sheet
	if (IS_GOOGLE) {
		var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
		  sheet = spreadSheet.getActiveSheet(),
		  range = sheet.getRange(2,1, artifacts.length, model.fields.length);
		
		range.setValues(artifactsAsCells);
		return JSON.parse(JSON.stringify(artifactsAsCells));
	} else {
	  return Excel.run({ delayForCellEdit: true }, function (context) {
		  var sheet = context.workbook.worksheets.getActiveWorksheet(),
			  range = sheet.getRangeByIndexes(1, 0, artifacts.length, model.fields.length);
		  range.values = artifactsAsCells;
		  return context.sync()
			  .then(function() {
				  return artifactsAsCells;
			  })
		})
	}
  }

// matches data against the fields to be shown in the spreadsheet - not all data fields are shown
// @param: array of the artifact objects we GOT from Spira
// @param: object of the meta information about the artifact
// @param: array of the fields that make up the sheet display
// @param: enum object of the different fieldTypes
// @param: array of the user objects
// @param: array of the component objects
// @param: array of the release objects
function matchArtifactsToFields(artifacts, artifactMeta, fields, fieldType, users, components, releases) {
  return artifacts.map(function(art) {
    return fields.map(function(field) {
      var originalFieldValue = "";
      
      // handle custom fields
      if (field.isCustom  && !art.isSubType) {   
        // if we have any custom props
        if (art.CustomProperties && art.CustomProperties.length) {
          // look for a match for the current field
          var customProp = art.CustomProperties.filter(function(custom) {
            return custom.Definition.CustomPropertyFieldName == field.field;
          });
          // if the property exists and isn't null - do a null check to handle booleans properly
          if (typeof customProp != "undefined" && customProp.length && customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]] !== null) {
            originalFieldValue = customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]];
          }
        }
       
      // handle subtype fields
      } else if (field.isSubTypeField) {
        if (artifactMeta.hasSubType && art.isSubType) {
          // first check to make sure the field exists in the artifact data
          if (typeof art[field.field] != "undefined" && art[field.field]) {
            originalFieldValue = art[field.field]
          }
        }
      
      // handle standard fields
      } else if (!art.isSubType) {
        // first check to make sure the field exists in the artifact data
        if (typeof art[field.field] != "undefined" && art[field.field]) {
          originalFieldValue = art[field.field];
        }
      }
      
      // handle list values - turn from the id to the actual string so the string can be displayed
      if (
        field.type == fieldType.drop || 
        field.type == fieldType.multi || 
        field.type == fieldType.user  || 
        field.type == fieldType.component  || 
        field.type == fieldType.release
      ) {
        return getListValueFromId(
          originalFieldValue, 
          field.type, 
          fieldType, 
          field.values, 
          users, 
          components, 
          releases
        );
      // handle date fields 
      } else if (field.type == fieldType.date) {
        return convertWcfDateToJs(originalFieldValue);
      // handle booleans - need to make sure null values are ignored ie treated differently to false
      } else if (field.type == fieldType.bool) {
        return originalFieldValue ? "Yes" : originalFieldValue === false ? "No" : "";
      // handle hierarchical artifacts
      } else if (field.setsHierarchy) {
        return makeHierarchical(originalFieldValue, art.IndentLevel);
      } else {
        return originalFieldValue;
      }
    });
  })
}

// takes an id for a lookup field and returns the string to display
// @param: int of the id to lookup
// @param: enum of the type of filed we need to look up
// @param: enum object of the different fieldTypes
// @param: array of the value objects for bespoke lookups
// @param: array of the user objects
// @param: array of the component objects
// @param: array of the release objects
function getListValueFromId(id, type, fieldType, fieldValues, users, components, releases) {
    var match = null;
    switch(type) {
      case fieldType.drop:
      case fieldType.multi:
        match = fieldValues.filter(function(val) { return val.id == id; });
        break;
      case fieldType.user:
        match = users.filter(function(val) { return val.id == id; });
        break;
      case fieldType.component:
        match = components.filter(function(val) { return val.id == id; });
        break;
      case fieldType.release:
        match = releases.filter(function(val) { return val.id == id; });
        break;
    }
    return typeof match != "undefined" && match && match.length ? match[0].name : "";
}

function makeHierarchical(value, indent) {
  var indentIncrements = Math.floor(indent.length/3) - 1;
  var indentText = "";
  for (var i = 0; i < indentIncrements; i++) {
    indentText += "> ";
  }
  indentText += value;
  return indentText;
}

// Parses a WCF JSON date in format /Date(1245398693390[-0500])/ to a JS date object without any timezone - to keep it in UTC
function convertWcfDateToJs(wcfDate) {
  var reDate = /^\/Date\((\d+)([+-]\d*)\)\//;
  if (reDate.test(wcfDate)) {
    var m = reDate.exec(wcfDate);
    var d = new Date();
    //We need to adjust the date to factor out the browser timezone
    var serverTicks = parseInt(m[1], 10);
    var serverOffset = 0;
        if (m.length >= 3)
        {
            var offset = Number(m[2]);
            var hoursOffset = Math.floor(offset / 100);
            var minsOffset = offset % 100;
            serverOffset = ((hoursOffset * 60) + minsOffset) * 60000;
        }
    d = new Date(serverTicks + serverOffset);
    return d;
  } else {
      //WCF bug where UTC dates come back already parsed!
      return wcfDate;
  }
}