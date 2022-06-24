/*
 *
 * =============
 * GENERAL SETUP
 * =============
 *
 */

// model becomes a new instance of the data store preserving the immutability of the primary data object.
var model = new Data();
var uiSelection = new tempDataStore();

// if devmode enabled, set the required fields and show the dev button
var devMode = true;
var isGoogle = typeof UrlFetchApp != "undefined";

/*
Global Variable to control if advanced options should be enabled to the user
Up to know, the advanced features are :
1."New Comment" field to all the artifacts -> allow creating new comments in Spira
2. Create new Artifacts Association:
  a. TestCase: Requirements, Releases and TestSet
  b. Requirements: Requirents
*/
var advancedMode = false;

//Global Variable to control if an admin advanced options button should be enabled to the logged user
var isAdmin = false;


//ENUMS

var UI_MODE = {
  initialState: 0,
  newProject: 1,
  newArtifact: 2,
  getData: 3,
  errorMode: 4
};

/*
 *
 * ============================
 * GOOGLE SHEETS SPECIFIC SETUP
 * ============================
 *
 */

// Google Sheets specific code to run at first launch
/*(function () {
  if (isGoogle) {

    // for dev mode only - comment out or set to false to disable any UI dev features
    setDevStuff(devMode);

    // add event listeners to the dom
    setEventListeners();

    // dom specific changes
    document.getElementById("help-connection-excel").style.display = "none";
  }
})();*/


/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 * Please comment/uncomment this block of code for Google Sheets/Excel
 */

import { params, templateFields, Data, tempDataStore } from './model.js';
import * as msOffice from './server.js';

export { showPanel, hidePanel };


// MS Excel specific code to run at first launch
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // on init make sure to run any required startup functions
    setEventListeners();
    // for dev mode only - comment out or set to false to disable any UI dev features
    setDevStuff(devMode);

    // dom specific changes
    document.body.classList.add('ms-office');
    document.getElementById("help-connection-google").style.display = "none";
  }
});


/* ==============================


/*
 *
 * =================================
 * UTILITIES & CROSS PANEL FUNCTIONS
 * =================================
 *
 */

function setDevStuff(devMode) {
  if (devMode) {
    document.getElementById("btn-dev").classList.remove("hidden");
    model.user.url = "https://internal-bruno.spiraservice.net/";
    model.user.userName = "administrator";
    model.user.api_key = btoa("&api-key=" + encodeURIComponent("{88C10992-22A6-47FD-B570-8A87624A8CBA}"));

    loginAttempt();
  }
}


function setEventListeners() {
  document.getElementById("btn-login").onclick = loginAttempt;
  document.getElementById("btn-help-login").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('login');
  };
  document.getElementById("lnk-help-login").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('login');
  };
  document.getElementById("btn-dev").onclick = setAuthDetails;
  document.getElementById("chkAdvanced").onclick = setAdvancedMode;


  document.getElementById("lnk-help-decide").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('modes')
  };
  document.getElementById("btn-decide-send").onclick = function () { showMainPanel("send") };
  document.getElementById("btn-decide-get").onclick = function () { showMainPanel("get") };
  document.getElementById("btn-decide-admin").onclick = function () { showAdminPanel() };
  document.getElementById("btn-decide-logout").onclick = logoutAttempt;
  document.getElementById("btn-help-main").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('data')
  };
  document.getElementById("btn-help-admin").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('login')
  };

  document.getElementById("btn-logout").onclick = logoutAttempt;
  document.getElementById("btn-logout-admin").onclick = logoutAttempt;
  document.getElementById("btn-main-back").onclick = hideMainPanel;
  document.getElementById("btn-admin-back").onclick = hideAdminPanel;

  // changing of dropdowns
  document.getElementById("select-product").onchange = changeProjectSelect;
  document.getElementById("select-artifact").onchange = changeArtifactSelect;

  document.getElementById("select-operation").onchange = changeOperationSelect;
  document.getElementById("select-template").onchange = changeTemplateSelect;

  document.getElementById("btn-toSpira").onclick = sendToSpiraAttempt;
  document.getElementById("btn-prepareTemplate").onclick = prepareTemplateAdmin;
  document.getElementById("btn-admin-send").onclick = sendToSpiraAttempt;

  document.getElementById("btn-fromSpira").onclick = getFromSpiraAttempt;
  document.getElementById("btn-adminGet").onclick = getFromSpiraAttempt;
  document.getElementById("btn-template").onclick = updateTemplateAttempt;
  document.getElementById("btn-updateToSpira").onclick = updateSpiraAttempt;

  document.getElementById("btn-help-back").onclick = function () { panelToggle("help") };
  document.getElementById("btn-help-section-login").onclick = function () { showChosenHelpSection('login') };
  document.getElementById("btn-help-section-modes").onclick = function () { showChosenHelpSection('modes') };
  document.getElementById("btn-help-section-data").onclick = function () { showChosenHelpSection('data') };
}



// used to show or hide / hide / show a specific panel
// @param: panel - string. suffix for items to act on (eg if id = panel-help, choice = "help")
function panelToggle(panel) {
  var panelElement = document.getElementById("panel-" + panel);
  panelElement.classList.toggle("offscreen");
}


function hidePanel(panel) {
  var panelElement = document.getElementById("panel-" + panel);
  panelElement.classList.add("offscreen");
}



function showPanel(panel) {
  var panelElement = document.getElementById("panel-" + panel);
  panelElement.classList.remove("offscreen");
}



// manage the loading spinner
function showLoadingSpinner() {
  document.getElementById("loader").classList.remove("hidden");
}



function hideLoadingSpinner() {
  document.getElementById("loader").classList.add("hidden");
}



// clear spreadsheet, model
function clearAddonData() {
  model = new Data();
  uiSelection = new tempDataStore();
  setDevStuff(devMode);
}



// clears the first sheet in the book
// @param: shouldClear - optional bool to check
function clearSheet(shouldClear) {
  var shouldClearToUse = typeof shouldClear !== 'undefined' ? shouldClear : true;
  if (shouldClearToUse) {
    if (isGoogle) {
      google.script.run
        .withSuccessHandler(newTemplateHandler)
        .clearAll(uiSelection);
      return true;
    } else {
      msOffice.clearAll(model)
        .then((response) => document.getElementById("panel-confirm").classList.add("offscreen"))
        .catch((error) => errorExcel(error));
    }
  }
  else { return false; }
}

//Handles the first step 
function newTemplateHandler(shouldContinue) {
  if (shouldContinue) {
    showLoadingSpinner();
    manageTemplateBtnState();

    // all data should already be loaded (as otherwise template button is disabled)
    // but check again that all data is present before kicking off template creation
    // if so, kicks off template creation, otherwise waits and tries again

    if (allGetsSucceeded()) {
      templateLoader();
      // otherwise, run an interval loop (should never get called as template button should be disabled)
    } else {
      var checkGetsSuccess = setInterval(attemptTemplateLoader, 1500);
      function attemptTemplateLoader() {
        if (allGetsSucceeded()) {
          templateLoader();
          clearInterval(checkGetsSuccess);
        }
      }
    }
  }
}

// resets the sidebar following logout
function resetSidebar() {
  // clear input field values
  document.getElementById("input-url").value = "";
  document.getElementById("input-userName").value = "";
  document.getElementById("input-password").value = "";

  // hide other panels, so login page is visible
  var otherPanels = document.querySelectorAll(".panel:not(#panel-auth)");
  // can't use forEach because that is not supported by Excel
  for (var i = 0; i < otherPanels.length; ++i) {
    otherPanels[i].classList.add("offscreen");
  }

  resetUi();
  // reset anything required if in devmode
  setDevStuff();
}

function resetUi() {
  // disable buttons and dropdowns
  document.getElementById("btn-template").disabled = true;
  document.getElementById("pnl-template").style.display = "none";
  document.getElementById("select-artifact").disabled = true;
  document.getElementById("btn-fromSpira").disabled = true;
  document.getElementById("btn-toSpira").disabled = true;
  document.getElementById("btn-updateToSpira").disabled = true;
  document.getElementById("btn-toSpira").innerHTML = "Prepare Template";

  // reset artifact dropdown to 'Select an Artifact'
  document.getElementById("select-artifact").selectedIndex = "0";
  // hide and clear the template info box
  document.getElementById("template-project").textContent = "";
  document.getElementById("template-artifact").textContent = "";

  // reset action buttons
  document.getElementById("btn-fromSpira").style.display = "";
  document.getElementById("btn-toSpira").style.display = "";
  document.getElementById("btn-updateToSpira").style.display = "";


  // reset guide text on the main pane
  document.getElementById("main-guide-1").classList.remove("pale");
  document.getElementById("main-guide-1-fromSpira").style.display = "";
  document.getElementById("main-guide-1-toSpira").style.display = "";
  document.getElementById("main-heading-fromSpira").style.display = "";
  document.getElementById("main-heading-toSpira").style.display = "";
  document.getElementById("main-guide-2").classList.add("pale");
  document.getElementById("main-guide-3").classList.add("pale");
}



// adds all options to a dropdown
// @param: selectId - is the id of the dom select element
// @param: array - the array of objects (with id, name, and optionally a disabled value, and hidden bool)
// @param: firstMessage - an optional text field to go at the top of the array - the initial choice 
function setDropdown(selectId, array, firstMessage) {
  // first make a deep copy of the array to stop any funny business
  var arrayCopy = JSON.parse(JSON.stringify(array)),
    select = document.getElementById(selectId);
  // if passed in, add default "select" option to top of project array
  if (firstMessage) arrayCopy.unshift({
    id: 0,
    name: firstMessage
  });
  // clear the dropdown
  select.innerHTML = "";
  arrayCopy.forEach(function (item) {
    var option = document.createElement("option");
    option.disabled = item.disabled;
    option.value = item.id;
    option.innerHTML = item.name;

    if (!item.hidden) {
      select.appendChild(option);
    }

  });
}


function isModelDifferentToSelection() {
  if (model.isTemplateLoaded) {
    if (!isAdmin) {
      var projectHasChanged = model.currentProject.id !== getSelectedProject().id;
      var artifactHasChanged = model.currentArtifact.id !== getSelectedArtifact().id;
      return projectHasChanged || artifactHasChanged;
    }
    else if (uiSelection.currentOperation == 2 || uiSelection.currentOperation == 3) {
      var templatetHasChanged = model.currentTemplate.id !== getSelectedTemplate().id;
      return templatetHasChanged;
    }
  } else {
    return false;
  }
}



/*
*
* ============
* LOGIN SCREEN
* ============
*
*/

// get user data from input fields and store in user data object
// adds the 'api-key' text before the key to make creating the urls simpler
function getAuthDetails() {
  model.user.url = document.getElementById("input-url").value;
  model.user.userName = document.getElementById("input-userName").value;
  var password = document.getElementById("input-password").value;
  model.user.api_key = btoa("&api-key=" + encodeURIComponent(password));
}



// fill in mock values for easy log in development, enable dev button
function setAuthDetails() {
  document.getElementById("input-url").value = model.user.url;
  document.getElementById("input-userName").value = model.user.userName;
  document.getElementById("input-password").value = model.user.password;
}


// switches the value of the global variable for the Advanced Mode
function setAdvancedMode() {
  if (document.getElementById('chkAdvanced').checked) {
    advancedMode = true;
  } else {
    advancedMode = false;
  }
}

// handle the click of the login button
function loginAttempt() {
  if (!devMode) getAuthDetails();
  login();
}

// login function that starts the intial data creation
function login() {
  artifactUpdateUI(UI_MODE.initialState);
  showLoadingSpinner();
  //First, check if this user is an admin (required to display admin options)
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(showHideAdminButton)
      .withFailureHandler(errorNetwork)
      .isUserAdmin(model.user);
  } else {
    msOffice.isUserAdmin(model.user)
      .then(response => showHideAdminButton(response.body))
      .catch(err => {
        return errorNetwork(err)
      }
      );
  }

  // call server side function to get projects
  // also serves as authentication check, if the user credentials aren't correct it will throw a network error
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(populateProjects)
      .withFailureHandler(errorNetwork)
      .getProjects(model.user);
  } else {
    msOffice.getProjects(model.user)
      .then(response => populateProjects(response.body))
      .catch(err => {
        return errorNetwork(err)
      }
      );
  }
}



// kick off prepping and showing main panel
// @param: projects - passed in projects data returned from the server following successful API call to Spira
function populateProjects(projects) {
  // take projects data from Spira API call, strip out unwanted fields, add to data model
  var pairedDownProjectsData = projects.map(function (project) {
    var result = {
      id: project.ProjectId,
      name: project.Name,
      templateId: project.ProjectTemplateId
    };
    return result;
  });

  // now add paired down project array to data store
  model.projects = pairedDownProjectsData;

  // sets the display current logged in user name
  document.getElementById("js--loggedInAs-decision").innerHTML = "Logged in as: " + model.user.userName;
  document.getElementById("js--loggedInAs-main").innerHTML = "Logged in as: " + model.user.userName;

  // get UI logic ready for decision panel
  showPanel("decide");
  hideLoadingSpinner();
}

// prepare prj template data to be displayed in the admin panel
// @param: templates - passed in templates data returned from the server following successful API call to Spira
function populateTemplates(templates) {
  // take projects data from Spira API call, strip out unwanted fields, add to data model
  var pairedDownTemplatesData = templates.reduce(function (filtered, template) {
    if (template.IsActive) {
      var result = {
        id: template.ProjectTemplateId,
        name: template.Name
      }
      filtered.push(result);
    }
    return filtered;
  }, []);
  // now add paired down template array to data store
  model.templates = pairedDownTemplatesData;
  //populates the templates dropdown menu
  setDropdown("select-template", model.templates, "Select a product template");
}


/*
*
* ===========
* MAIN SCREEN
* ===========
*
*/

// manage the switching of the UI off the login screen on succesful login and retrieval of projects
function showMainPanel(type) {

  //all the users should have access to these options
  isAdmin = false;

  setDropdown("select-product", model.projects, "Select a product");
  if (type == "get") {
    var paramsGet = params.artifacts.filter((element) => {
      if (!element.hasOwnProperty("sendOnly")) {
        return element;
      }
    })

    setDropdown("select-artifact", paramsGet, "Select an artifact");
  }
  else {
    setDropdown("select-artifact", params.artifacts, "Select an artifact");
  }




  // set the buttons to the correct mode
  if (type == "send") {
    document.getElementById("btn-fromSpira").style.display = "none";
    document.getElementById("main-guide-1-fromSpira").style.display = "none";
    document.getElementById("main-heading-fromSpira").style.display = "none";
    document.getElementById("btn-updateToSpira").style.visibility = "hidden";
    document.getElementById("main-guide-3").style.visibility = "hidden";
  } else {
    document.getElementById("btn-toSpira").style.display = "none";
    document.getElementById("main-guide-1-toSpira").style.display = "none";
    document.getElementById("main-heading-toSpira").style.display = "none";
    document.getElementById("main-guide-3").style.visibility = "visible";
    document.getElementById("btn-updateToSpira").style.visibility = "visible";
  }

  // opens the panel
  showPanel("main");
  hideLoadingSpinner();
}


function hideMainPanel() {
  hidePanel("main");
  // reset the buttons and dropdowns
  resetUi();
  uiSelection = new tempDataStore();
  // make sure the system does not think any data is loaded
  model.isTemplateLoaded = false;
}


// run server side code to manage logout
function logoutAttempt() {
  var message = 'All data on the active sheet will be deleted. Continue?'
  //warn user that all data on the first sheet will be lost. Returns true or false
  showPanel("confirm");
  document.getElementById("message-confirm").innerHTML = message;
  document.getElementById("btn-confirm-ok").onclick = () => logout(true);
  document.getElementById("btn-confirm-cancel").onclick = () => hidePanel("confirm");
}



// @param: shouldLogout - a true or false value from google/Excel
function logout(shouldLogout) {
  if (shouldLogout) {
    clearAddonData();
    resetSidebar();
  }
}



function changeProjectSelect(e) {
  //sets the UI to correspond to this mode
  artifactUpdateUI(UI_MODE.newProject);

  // if the project field has not been selected all other selected buttons are disabled
  if (e.target.value == 0) {
    document.getElementById("select-artifact").disabled = true;
    document.getElementById("btn-toSpira").disabled = true;
    document.getElementById("btn-fromSpira").disabled = true;
    document.getElementById("btn-template").disabled = true;
    uiSelection.currentProject = null;
  } else {
    // enable artifacts dropdown
    document.getElementById("select-artifact").disabled = false;

    // get the project object and update project information if project has changed
    var chosenProject = getSelectedProject();
    if (chosenProject.id && chosenProject.id !== uiSelection.currentProject.id) {
      //set the temp data store project to the one selected;
      uiSelection.currentProject = chosenProject;

      // enable template button only when all info is received - otherwise keep it disabled
      manageTemplateBtnState();

      // kick off API calls
      getProjectSpecificInformation(model.user, uiSelection.currentProject.id);

      // for 6.1 the v6 API for get projects does not get the project template IDs so have to do this
      getTemplateFromProjectId(model.user, uiSelection.currentProject.id, uiSelection.currentArtifact);

      // get new data for artifact and this project, if artifact has been selected
      // USE THIS CODE WHEN bug in 6.1 is fixed
      // if (uiSelection.currentArtifact) {
      //getArtifactSpecificInformation(model.user, uiSelection.currentProject.templateId, uiSelection.currentProject.id, uiSelection.currentArtifact)
      // }
    }
  }
}

function getTemplateFromProjectId(user, projectId, artifact) {
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getArtifactSpecificInformationInterim)
      .withFailureHandler(errorNetwork)
      .getTemplateFromProjectId(user, projectId);
  } else {
    msOffice.getTemplateFromProjectId(user, projectId)
      .then((response) => getArtifactSpecificInformationInterim(response.body))
      .catch((error) => errorNetwork(error));
  }

  function getArtifactSpecificInformationInterim(template) {
    uiSelection.currentProject.templateId = template.ProjectTemplateId;
    if (uiSelection.currentArtifact) {
      getArtifactSpecificInformation(model.user, template.ProjectTemplateId, uiSelection.currentProject.id, uiSelection.currentArtifact)
    }
  }
}

//handles hiding/displaying and changing colors of elements in the UI based on the operation
function artifactUpdateUI(mode) {

  switch (mode) {

    case UI_MODE.initialState:
      //when re-starting session

      document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'bold';
      document.getElementById("main-guide-1").classList.remove("pale");

      document.getElementById('main-guide-2').style.fontWeight = 'normal';
      document.getElementById("main-guide-2").classList.add("pale");
      document.getElementById("btn-fromSpira").disabled = true;

      document.getElementById('main-guide-3').style.fontWeight = 'normal';
      document.getElementById("btn-updateToSpira").disabled = true;

      document.getElementById('btn-fromSpira').classList.remove('ms-Button--default');
      document.getElementById('btn-fromSpira').classList.add('ms-Button--primary');

      break;

    case UI_MODE.newProject:
      //when selecting a new project
      document.getElementById("main-guide-1").classList.remove("pale");
      document.getElementById("main-guide-2").classList.add("pale");
      document.getElementById("main-guide-3").classList.add("pale");
      document.getElementById("btn-fromSpira").disabled = true;
      document.getElementById("btn-updateToSpira").disabled = true;
      break;

    case UI_MODE.newArtifact:
      //when selecting a new artifact
      document.getElementById("main-guide-3").classList.add("pale");
      document.getElementById("btn-updateToSpira").disabled = true;
      break;

    case UI_MODE.getData:
      //when clicking from-Spira button - the admin pane is different from the main pane
      if (isAdmin === false) {
        document.getElementById("main-guide-admin-2-get").disabled = false;
      }
      else {
        document.getElementById('main-guide-admin-3-post').style.fontWeight = 'normal';
        document.getElementById("main-guide-admin-3-post").classList.remove("pale");
        document.getElementById('main-guide-admin-3-post').style.fontWeight = 'bold';
        document.getElementById("btn-admin-send").disabled = false;
      }
      break;

    case UI_MODE.errorMode:
      //in case of any error
      document.getElementById("main-guide-2").classList.remove("pale");
      document.getElementById("btn-fromSpira").disabled = false;
      document.getElementById("main-guide-3").classList.add("pale");
      document.getElementById("btn-updateToSpira").disabled = true;
      break;
  }
}


function changeArtifactSelect(e) {
  //sets the UI to correspond to this mode
  artifactUpdateUI(UI_MODE.newArtifact);
  if (e.target.value == 0) {
    document.getElementById("btn-toSpira").disabled = true;
    document.getElementById("btn-fromSpira").disabled = true;
    document.getElementById("btn-template").disabled = true;
    uiSelection.currentArtifact = null;
  } else {
    // get the artifact object and update artifact information if artifact has changed
    var chosenArtifact = getSelectedArtifact();

    if (chosenArtifact !== uiSelection.currentArtifact) {
      //set the temp date store artifact to the one selected;
      uiSelection.currentArtifact = chosenArtifact;

      // enable template button only when all info is received - otherwise keep it disabled
      manageTemplateBtnState();
      // kick off API calls - if we have a current template and project
      if (uiSelection.currentProject.templateId && uiSelection.currentProject.id) {
        getArtifactSpecificInformation(model.user, uiSelection.currentProject.templateId, uiSelection.currentProject.id, uiSelection.currentArtifact);
      }
    }
  }
}



// disables and enables the main action buttons based on status of required API calls
function manageTemplateBtnState() {
  // initially disable the button, because required API calls not completed
  document.getElementById("btn-toSpira").disabled = true;
  document.getElementById("btn-fromSpira").disabled = true;
  document.getElementById("btn-template").disabled = true;

  // only try to enable the button when both a project and artifact have been chosen
  if (uiSelection.currentProject && uiSelection.currentArtifact) {
    // set a function to run repeatedly until all gets are done
    // then enable the button, and stop the timer loop
    var checkGetsSuccess = setInterval(updateButtonStatus, 500);

    // and show a message while api calls are underway
    document.getElementById("message-fetching-data").style.visibility = "visible";

    function updateButtonStatus() {
      if (allGetsSucceeded()) {

        if (!document.getElementById("btn-updateToSpira").disabled) {
          //Send to Spira is active - click on Get from Spira
          //sets the UI to allow update
          document.getElementById("btn-fromSpira").disabled = false;

          document.getElementById("main-guide-2").classList.add("pale");
          document.getElementById("main-guide-3").classList.remove("pale");
          document.getElementById("message-fetching-data").style.visibility = "hidden";
        }
        else {
          //Send to Spira is NOT active - project is selected
          document.getElementById("btn-toSpira").disabled = false;
          document.getElementById("btn-fromSpira").disabled = false;
          document.getElementById("btn-updateToSpira").disabled = true;

          document.getElementById('btn-fromSpira').classList.remove('ms-Button--default');
          document.getElementById('btn-fromSpira').classList.add('ms-Button--primary');

          document.getElementById("message-fetching-data").style.visibility = "hidden";

          document.getElementById("main-guide-1").classList.add("pale");
          document.getElementById("main-guide-2").classList.remove("pale");
          document.getElementById("main-guide-3").classList.add("pale");

          document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'normal';
          document.getElementById('main-guide-2').style.fontWeight = 'bold';
          document.getElementById('main-guide-3').style.fontWeight = 'normal';
        }

        clearInterval(checkGetsSuccess);

        // if there is a discrepancy between the dropdown and the currently active template
        // only do this is user will send to spira - determined by whether the send to spira button is visible
        if (document.getElementById("btn-toSpira").style.display != "none") {
          if (isModelDifferentToSelection(false)) {
            document.getElementById("pnl-template").style.display = "";
            document.getElementById("btn-template").disabled = false;
          } else {
            document.getElementById("pnl-template").style.display = "none";
            document.getElementById("btn-template").disabled = true;
          }
        }
      }
      else {
      }
    }
  }
}



// starts the process to create a template from chosen options
function createTemplateAttempt() {
  var message = 'The active sheet will be replaced. Continue?'
  //warn the user data will be erased
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(clearSheet)
      .warn(message);
  } else {
    showPanel("confirm");
    document.getElementById("message-confirm").innerHTML = message;
    document.getElementById("btn-confirm-ok").onclick = () => createTemplate(true);
    document.getElementById("btn-confirm-cancel").onclick = () => hidePanel("confirm");
  }
}



//@param: shouldClearForm - boolean response from confirmation dialog above
function createTemplate(shouldContinue) {
  if (shouldContinue) {
    clearSheet();
    showLoadingSpinner();
    manageTemplateBtnState();

    // all data should already be loaded (as otherwise template button is disabled)
    // but check again that all data is present before kicking off template creation
    // if so, kicks off template creation, otherwise waits and tries again
    // the exception is when using advanced admin mode operations not based on projects
    if (allGetsSucceeded() || uiSelection.currentOperation == 1 || uiSelection.currentOperation == 2 || uiSelection.currentOperation == 3) {
      templateLoader();
      // otherwise, run an interval loop (should never get called as template button should be disabled)
    } else {
      var checkGetsSuccess = setInterval(attemptTemplateLoader, 500);
      function attemptTemplateLoader() {
        if (allGetsSucceeded()) {
          templateLoader();
          clearInterval(checkGetsSuccess);
        }
      }
    }
  }
}



function getFromSpiraAttempt() {
  // first update state to reflect user intent
  model.isGettingDataAttempt = true;
  //check that template is loaded and that it matches the UI choices
  if (model.isTemplateLoaded && !isModelDifferentToSelection()) {
    showLoadingSpinner();
    //call export function
    if (isGoogle) {
      google.script.run
        .withFailureHandler(errorImpExp)
        .withSuccessHandler(getFromSpiraComplete)
        .getFromSpiraGoogle(model, params.fieldType, advancedMode);
    } else {
      msOffice.getFromSpiraExcel(model, params.fieldType)
        .then((response) => getFromSpiraComplete(response))
        .catch((error) => errorImpExp(error));
    }
  } else {
    //if no template - then get the template
    createTemplateAttempt();
  }
  //sets the UI to correspond to this mode
  artifactUpdateUI(UI_MODE.getData);
}


function getFromSpiraComplete(log) {
  if (devMode)
    //if array (which holds error responses) is present, and errors present
    if (log && log.errorCount) {
      errorMessages = log.entries
        .filter(function (entry) { return entry.error; })
        .map(function (entry) { return entry.message; });
    }
    else {
      manageTemplateBtnState();
    }
  hideLoadingSpinner();

  //runs the export success function, passes a boolean flag, if there are errors the flag is true.
  if (log && log.status) {
    if (isGoogle) {
      google.script.run.operationComplete(log.status);
    } else {
      msOffice.operationComplete(log.status);
    }
  }
}



function sendToSpiraAttempt() {
  // first update state to reflect user intent
  model.isGettingDataAttempt = false;
  //check that template is loaded
  if (model.isTemplateLoaded) {
    showLoadingSpinner();

    //call export function
    if (isGoogle) {
      google.script.run
        .withFailureHandler(errorImpExp)
        .withSuccessHandler(sendToSpiraComplete)
        .sendToSpira(model, params.fieldType, false);
    } else {
      msOffice.sendToSpira(model, params.fieldType, false)
        .then((response) => sendToSpiraComplete(response))
        .catch((error) => errorImpExp(error));
    }
  } else {
    //if no template - then get the template
    createTemplateAttempt();
  }
}

function prepareTemplateAdmin() {
  if (!model.isTemplateLoaded) {
    createTemplateAttempt();
    document.getElementById("btn-prepareTemplate").disabled = true;
    document.getElementById("main-guide-admin-2-send").classList.add("pale");
    document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'normal';

    document.getElementById("main-guide-admin-3-post").classList.remove("pale");
    document.getElementById("btn-admin-send").classList.remove("pale");
    document.getElementById('main-guide-admin-3-post').style.fontWeight = 'bold';

    document.getElementById("btn-admin-send").disabled = false;

  }
}

function updateSpiraAttempt() {

  // first update state to reflect user intent
  model.isGettingDataAttempt = false;
  //check that template is loaded
  if (model.isTemplateLoaded) {
    showLoadingSpinner();

    //call export function
    if (isGoogle) {
      google.script.run
        .withFailureHandler(errorImpExp)
        .withSuccessHandler(sendToSpiraComplete)
        .sendToSpira(model, params.fieldType, true);
    } else {
      msOffice.sendToSpira(model, params.fieldType, true)
        .then((response) => sendToSpiraComplete(response))
        .catch((error) => errorImpExp(error));
    }
  } else {
    //if no template - throw an error
    errorExcel("The spreadsheet does not match the selected artifact. Please check your data.")
  }

}

function sendToSpiraComplete(log) {

  if (isGoogle && log) {
    log = JSON.parse(log);
  }
  hideLoadingSpinner();
  if (devMode) console.log(log);

  //if array (which holds error responses) is present, and errors present
  if (log.errorCount) {
    var errorMessages = log.entries
      .filter(function (entry) { return entry.error; })
      .map(function (entry) { return entry.message; });

  }
  //runs the export success function, passes a boolean flag, if there are errors the flag is true.
  if (log && log.status) {
    if (isGoogle) {
      google.script.run.operationComplete(log.status);
    } else {
      msOffice.operationComplete(log.status);
    }
  }

}



function updateTemplateAttempt() {
  // first update state to reflect user intent
  model.isGettingDataAttempt = false;
  createTemplateAttempt();
}








/*
*
* ===========
* HELP SCREEN
* ===========
*
*/
// manage showing the correct help section to the user
// @param: choice - string. suffix for items to select (eg if id = help-section-fields, choice = "fields")
function showChosenHelpSection(choice) {
  // does not use a dynamic list using queryselectorall and node list because Excel does not support this
  // hide all sections and then only show the one the user wants
  document.getElementById("help-section-login").classList.add("hidden");
  document.getElementById("help-section-modes").classList.add("hidden");
  document.getElementById("help-section-data").classList.add("hidden");
  document.getElementById("help-section-" + choice).classList.remove("hidden");

  // set all buttons back to normal, then highlight one just clicked
  document.getElementById("btn-help-section-login").classList.remove("create");
  document.getElementById("btn-help-section-modes").classList.remove("create");
  document.getElementById("btn-help-section-data").classList.remove("create");
  document.getElementById("btn-help-section-" + choice).classList.add("create");
}



/*
*
* ===========
* ADMIN SCREEN
* ===========
*
*/
// Admin-exclusive: retrieves the operation object that matches the one selected in the dropdown
// returns a operation object
function getSelectedOperation() {
  // store dropdown value
  var select = document.getElementById("select-operation");
  var operationDropdownVal = select.options[select.selectedIndex].value;
  // filter the project lists to those chosen 
  var operationSelected = model.operations.filter(function (operation) {
    return operation.id == operationDropdownVal;
  })[0];
  return operationSelected;
}

//decide if the administrator advanced options button should be displayed for the current logged user
function showHideAdminButton(spiraUser) {
  spiraUser.Admin ? document.getElementById("btn-decide-admin").style.visibility = "visible" : document.getElementById("btn-decide-admin").style.visibility = "hidden";
}


// manage the switching of the UI off the login screen to the administrator screen 
function showAdminPanel() {
  //menu exclusive for system administrator users
  isAdmin = true;

  //hide information we don't want to be displayed yet
  document.getElementById("main-guide-admin-2-get").style.visibility = "hidden";
  document.getElementById("main-guide-admin-2-send").style.visibility = "hidden";
  document.getElementById("btn-prepareTemplate").style.visibility = "hidden";
  document.getElementById("btn-adminGet").style.visibility = "hidden";
  document.getElementById("main-guide-admin-3-put").style.visibility = "hidden";
  document.getElementById("main-guide-admin-3-post").style.visibility = "hidden";
  document.getElementById("btn-admin-send").style.visibility = "hidden";
  document.getElementById("main-guide-admin-templates").style.visibility = "hidden";
  document.getElementById("select-template").style.visibility = "hidden";

  document.getElementById("btn-prepareTemplate").disabled = true;
  document.getElementById("main-guide-admin-2-send").classList.add("pale");
  document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'bold';

  document.getElementById("main-guide-admin-3-post").classList.add("pale");
  document.getElementById("btn-admin-send").classList.add("pale");
  document.getElementById('main-guide-admin-3-post').style.fontWeight = 'normal';

  document.getElementById("btn-admin-send").disabled = false;

  //Get the project templates now we know this user is an admin
  // call server side function to get templates
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(populateTemplates)
      .withFailureHandler(errorNetwork)
      .getProjectTemplates(model.user);
  } else {
    msOffice.getProjectTemplates(model.user)
      .then(response => populateTemplates(response.body))
      .catch(err => {
        return errorNetwork(err)
      }
      );
  }

  //populate the operations dropdown menu
  setDropdown("select-operation", model.operations, "Select an operation");

  showPanel("admin");
  hideLoadingSpinner();

}

function hideAdminPanel() {
  hidePanel("admin");
  // reset the buttons and dropdowns
  resetUi();
  uiSelection = new tempDataStore();
  // make sure the system does not think any data is loaded
  model.isTemplateLoaded = false;
}

function changeOperationSelect(e) {

  // if the operation field has not been selected all other objects in this panel will be hidden
  if (e.target.value == 0) {
    //hide information we don't want to be displayed
    document.getElementById("main-guide-admin-2-get").style.visibility = "hidden";
    document.getElementById("main-guide-admin-2-send").style.visibility = "hidden";
    document.getElementById("btn-prepareTemplate").style.visibility = "hidden";
    document.getElementById("btn-adminGet").style.visibility = "hidden";
    document.getElementById("main-guide-admin-3-put").style.visibility = "hidden";
    document.getElementById("main-guide-admin-3-post").style.visibility = "hidden";
    document.getElementById("btn-admin-send").style.visibility = "hidden";
    document.getElementById("main-guide-admin-templates").style.visibility = "hidden";
    document.getElementById("select-template").style.visibility = "hidden";

    document.getElementById("btn-prepareTemplate").disabled = true;
    document.getElementById("main-guide-admin-2-send").classList.add("pale");
    document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'normal';

    document.getElementById("main-guide-admin-3-post").classList.add("pale");
    document.getElementById("btn-admin-send").classList.add("pale");
    document.getElementById('main-guide-admin-3-post').style.fontWeight = 'normal';

    document.getElementById("btn-admin-send").disabled = false;

    uiSelection.currentOperation = null;

  } else {
    // enable other objects, depending on the oparation selected
    var chosenOperation = getSelectedOperation();

    switch (chosenOperation.type) {
      case "send-system":
        //system wide operations that requires send data only
        if (chosenOperation.id == 1) {
          //Create users operation
          //Display the necessary objects on the taskpane
          document.getElementById("main-guide-admin-2-get").style.display = "none";
          document.getElementById("main-guide-admin-2-get").style.visibility = "hidden";

          document.getElementById("main-guide-admin-2-send").classList.remove("pale");
          document.getElementById("main-guide-admin-2-send").style.visibility = "visible";
          document.getElementById("main-guide-admin-2-send").style.display = "";

          document.getElementById("btn-prepareTemplate").classList.remove("pale");
          document.getElementById("btn-prepareTemplate").style.visibility = "visible";
          document.getElementById("btn-prepareTemplate").disabled = false;
          document.getElementById("btn-prepareTemplate").style.display = "";

          document.getElementById("select-template").style.display = "";

          document.getElementById("btn-adminGet").style.display = "none";
          document.getElementById("btn-adminGet").style.visibility = "hidden";

          document.getElementById("main-guide-admin-3-post").style.visibility = "visible";

          document.getElementById("main-guide-admin-3-put").style.display = "none";

          document.getElementById("btn-admin-send").style.visibility = "visible";

          document.getElementById("main-guide-admin-templates").style.visibility = "hidden";
          document.getElementById("main-guide-admin-templates").style.display = "none";
          document.getElementById("select-template").style.display = "none";

          uiSelection.currentOperation = 1;
          //sets the selected artifact based on admin operation
          uiSelection.currentArtifact = getAdminArtifact();
          //get bespoke fields for this operation's artifact
          getArtifactSpecificInformation(model.user, null, null, uiSelection.currentArtifact);
        }
        break;

      case "send-template":
        //template-based operations that requires send data only
        if (chosenOperation.id == 2) {
          //Create custom lists and values operation
          //Display the necessary objects on the taskpane
          document.getElementById("main-guide-admin-2-get").style.display = "none";

          document.getElementById("main-guide-admin-2-send").classList.remove("pale");
          document.getElementById("main-guide-admin-2-send").style.visibility = "visible";
          document.getElementById("main-guide-admin-2-send").style.display = "";

          document.getElementById("btn-prepareTemplate").classList.remove("pale");
          document.getElementById("btn-prepareTemplate").style.visibility = "visible";
          document.getElementById("btn-prepareTemplate").disabled = true;
          document.getElementById("btn-prepareTemplate").style.display = "";


          document.getElementById("btn-adminGet").style.display = "none";
          document.getElementById("btn-adminGet").style.visibility = "hidden";

          document.getElementById("main-guide-admin-3-post").style.visibility = "visible";

          document.getElementById("main-guide-admin-3-put").style.display = "none";

          document.getElementById("btn-admin-send").style.visibility = "visible";

          document.getElementById("main-guide-admin-templates").style.visibility = "visible";
          document.getElementById("select-template").style.visibility = "visible";
          document.getElementById("main-guide-admin-templates").style.display = "";
          document.getElementById("select-template").style.display = "";
          document.getElementById("select-template").selectedIndex = 0;

          uiSelection.currentOperation = 2;
          //sets the selected artifact based on admin operation
          uiSelection.currentArtifact = getAdminArtifact();
        }
        break;

      case "get-template":
        //template-based operations that requires get data + send later
        if (chosenOperation.id == 3) {
          //Create custom lists and values operation
          //Display the necessary objects on the taskpane
          document.getElementById("main-guide-admin-2-send").style.display = "none";
          document.getElementById("main-guide-admin-2-send").style.visibility = "hidden";

          document.getElementById("main-guide-admin-2-get").classList.remove("pale");
          document.getElementById("main-guide-admin-2-get").style.visibility = "visible";
          document.getElementById("main-guide-admin-2-get").style.display = "";

          document.getElementById("btn-adminGet").classList.remove("pale");
          document.getElementById("btn-adminGet").style.visibility = "visible";
          document.getElementById("btn-adminGet").disabled = true;

          document.getElementById("btn-adminGet").style.display = "";

          document.getElementById("btn-prepareTemplate").style.visibility = "hidden";
          document.getElementById("btn-prepareTemplate").style.display = "none";

          document.getElementById("main-guide-admin-3-post").style.visibility = "visible";

          document.getElementById("main-guide-admin-3-put").style.display = "none";

          document.getElementById("btn-admin-send").style.visibility = "visible";

          document.getElementById("main-guide-admin-templates").style.visibility = "visible";
          document.getElementById("select-template").style.visibility = "visible";
          document.getElementById("main-guide-admin-templates").style.display = "";

          document.getElementById("select-template").style.display = "";
          document.getElementById("select-template").selectedIndex = 0;

          uiSelection.currentOperation = 3;
          //sets the selected artifact based on admin operation
          uiSelection.currentArtifact = getAdminArtifact();
        }
        break;
    }

  }
  model.isTemplateLoaded = false;
}


function changeTemplateSelect(e) {
  //First, enable the respective operation next button
  if (e.target.value != 0) {
    if (uiSelection.currentOperation == 2) {
      document.getElementById("btn-prepareTemplate").disabled = false;
    } else if (uiSelection.currentOperation == 3) {
      document.getElementById("btn-adminGet").disabled = false;
    }
    var chosenTemplate = getSelectedTemplate();
    if (chosenTemplate.id && chosenTemplate.id !== uiSelection.currentTemplate.id) {
      //set the temp data store project to the one selected;
      uiSelection.currentTemplate = chosenTemplate;
    }
  }
  else {
    document.getElementById("btn-prepareTemplate").disabled = true;
    document.getElementById("btn-adminGet").disabled = true;
    uiSelection.currentTemplate = null;
  }
}

//returns the artifactId associated with the given admin operation
function getAdminArtifact() {
  //filter the current operation
  var operationSelected = model.operations.filter(function (operation) {
    return operation.id == uiSelection.currentOperation;
  })[0];

  // filter the artifact lists to those chosen 
  var artifactSelected = params.artifacts.filter(function (artifact) {
    return artifact.id == operationSelected.artifactId;
  })[0];
  return artifactSelected;
}

/*
*
* =================
* CREATING TEMPLATE
* =================
*
*/

// retrieves the project object that matches the project selected in the dropdown
// returns a project object
function getSelectedProject() {
  // store dropdown value
  var select = document.getElementById("select-product");
  var projectDropdownVal = select.options[select.selectedIndex].value;
  // filter the project lists to those chosen 
  var projectSelected = model.projects.filter(function (project) {
    return project.id == projectDropdownVal;
  })[0];
  return projectSelected;
}



// retrieves the artifact object that matches the artifact selected in the dropdown
// returns an artifact object
function getSelectedArtifact() {
  // store dropdown values
  var select = document.getElementById("select-artifact");
  var artifactDropdownVal = select.options[select.selectedIndex].value;

  // filter the artifact lists to those chosen 
  var artifactSelected = params.artifacts.filter(function (artifact) {
    return artifact.id == artifactDropdownVal;
  })[0];
  return artifactSelected;
}

// retrieves the template object that matches the template selected in the administration screen dropdown
// returns a template object
function getSelectedTemplate() {
  // store dropdown value
  var select = document.getElementById("select-template");
  var projectDropdownVal = select.options[select.selectedIndex].value;
  // filter the project lists to those chosen 
  var templateSelected = model.templates.filter(function (template) {
    return template.id == projectDropdownVal;
  })[0];
  return templateSelected;
}

// kicks off all relevant API calls to get project specific information
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getProjectSpecificInformation(user, projectId) {
  model.projectGetRequestsMade = 0;
  // get project information
  getUsers(user, projectId);
  getComponents(user, projectId);
  getReleases(user, projectId);
}



// kicks off all relevant API calls to get artifact specific information
// @param: user - the user object of the logged in user
// @param: templateId - int of the reqested project template
// @param: artifact - object of the reqested artifact - needed to query on different parts of object
function getArtifactSpecificInformation(user, templateId, projectId, artifact) {

  // first reset get counts
  model.artifactGetRequestsMade = 0;
  model.artifactGetRequestsToMake = model.baselineArtifactGetRequests;
  // increase the count if any bespoke fields are present (eg folders or incident types)
  var bespokeData = fieldsWithBespokeData(templateFields[artifact.field]);

  if (bespokeData) {
    model.artifactGetRequestsToMake += bespokeData.length;
    // get any bespoke field information
    bespokeData.forEach(function (bespokeField) {
      getBespoke(user, templateId, projectId, artifact.field, bespokeField);
    });
  }
  // get standard artifact information - eg custom fields
  if (templateId != null && projectId != null) {
    getCustoms(user, templateId, artifact.id);
  }
}



// goes through artifact object and returns an array of field objects that have specific rest calls to get their data
// @param: artifact - object of the requested artifact
function fieldsWithBespokeData(artifactFields) {
  if (!artifactFields.length) {
    return;
  }
  var bespokeFields = artifactFields.filter(function (field) {
    return field.bespoke;
  });
  return bespokeFields.length ? bespokeFields : false;
}



// starts GET request to Spira for project components
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getComponents(user, projectId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getComponentsSuccess)
      .withFailureHandler(errorNetwork)
      .getComponents(user, projectId);
  } else {
    msOffice.getComponents(user, projectId)
      .then((response) => getComponentsSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets component data on the model
function getComponentsSuccess(data) {
  // clear old values
  uiSelection.projectComponents = [];
  // add relevant data to the main model store
  uiSelection.projectComponents = data
    .filter(function (item) { return item.IsActive; })
    .map(function (item) {
      return {
        id: item.ComponentId,
        name: item.Name
      };
    });
  model.projectGetRequestsMade++;
}



// starts GET request to Spira for project / artifact custom properties
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
// @param: artifactId - int of the reqested artifact
function getCustoms(user, projectId, artifactId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getCustomsSuccess)
      .withFailureHandler(errorNetwork)
      .getCustoms(user, projectId, artifactId);
  } else {
    msOffice.getCustoms(user, projectId, artifactId)
      .then((response) => getCustomsSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets custom field data on the model - adding to a temp holding area, to allow for changes before template creation
function getCustomsSuccess(data) {
  // clear old values
  uiSelection.artifactCustomFields = [];
  // assign unparsed data to data object
  // these values are parsed later depending on function needs
  uiSelection.artifactCustomFields = data
    .filter(function (item) { return !item.IsDeleted; })
    .map(function (item) {
      var customField = {
        isCustom: true,
        field: item.CustomPropertyFieldName,
        name: item.Name,
        propertyNumber: item.PropertyNumber,
        type: item.CustomPropertyTypeId,
      };

      // mark as required or not - default is that it can be empty
      var allowEmptyOption = item.Options && item.Options.filter(function (option) {
        return option.CustomPropertyOptionId && option.CustomPropertyOptionId === 1;
      });
      if (allowEmptyOption && allowEmptyOption.length && allowEmptyOption[0].Value == "N") {
        customField.required = true;
      }
      // add array of values for dropdowns
      if (item.CustomPropertyTypeId == params.fieldType.drop || item.CustomPropertyTypeId == params.fieldType.multi) {
        customField.values = item.CustomList.Values.map(function (listItem) {
          return {
            id: listItem.CustomPropertyValueId,
            name: listItem.Name
          };
        });
      }
      return customField;
    }
    );
  model.artifactGetRequestsMade++;
}



// starts GET request to Spira for project users properties
// @param: user - the user object of the logged in user
// @param: templateId - int of the reqested template
function getBespoke(user, templateId, projectId, artifactId, field) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getBespokeSuccess)
      .withFailureHandler(errorNetwork)
      .getBespoke(user, templateId, projectId, artifactId, field);
  } else {
    msOffice.getBespoke(user, templateId, projectId, artifactId, field)
      .then((response) => getBespokeSuccess({
        artifactName: artifactId,
        field: field,
        values: response.body
      }))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets user data on the model
function getBespokeSuccess(data) {
  // create and clear old values
  if (typeof uiSelection[data.artifactName] == "undefined") {
    uiSelection[data.artifactName] = {};
  }
  uiSelection[data.artifactName][data.field.field] = [];

  // if there is data take steps to add it to the artifact object
  if (data && data.values && data.values.length) {
    // map through user obj and assign names
    var values = data.values.map(function (item) {
      var obj = {};
      obj.id = item[data.field.bespoke.idField];
      obj.name = item[data.field.bespoke.nameField];
      if (data.field.bespoke.indent) {
        obj.indent = item[data.field.bespoke.indent];
      }
      return obj;
    });

    // indented fields need to specify a field name that contains indent data, so we use this as a check to see if a field is hierarchicel
    if (data.field.bespoke.indent) {
      var hierarchicalValues = values.sort(function (a, b) {
        if (a.indent < b.indent) {
          return -1;
        }
        if (a.indent > b.indent) {
          return 1;
        }
        // names must be equal
        return 0;
      }).map(function (x) {
        var indentAmount = (x.indent.length / 3) - 1;
        var indentString = "...";
        x.name = (indentString.repeat(indentAmount)) + x.name;
        return x;
      });
      uiSelection[data.artifactName][data.field.field].values = hierarchicalValues;

    } else {
      uiSelection[data.artifactName][data.field.field].values = values;
    }
  }
  // in all cases make sure the successful request is recorded
  model.artifactGetRequestsMade++;
}



// starts GET request to Spira for project releases properties
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getReleases(user, projectId, artifactId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getReleasesSuccess)
      .withFailureHandler(errorNetwork)
      .getReleases(user, projectId);
  } else {
    msOffice.getReleases(user, projectId)
      .then((response) => getReleasesSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}
// formats and sets release data on the model
function getReleasesSuccess(data) {
  //Getting the Active releases (for standard Release fields)
  // clear old values
  uiSelection.projectActiveReleases = [];
  // add relevant data to the main model store
  var activeReleases = data.map(function (item) {
    //getting only the active releases
    if (item.Active) {
      return {
        id: item.ReleaseId,
        name: item.Name
      };
    }
  });

  uiSelection.projectActiveReleases = activeReleases.filter(function (item) {
    if (typeof item !== "undefined") {
      return item;
    }
  });

  //Getting all the releases in the project (for custom Release fields)
  // clear old values
  uiSelection.projectReleases = [];
  // add relevant data to the main model store
  uiSelection.projectReleases = data.map(function (item) {
    return {
      id: item.ReleaseId,
      name: item.Name
    };
  });

  model.projectGetRequestsMade++;
}



// starts GET request to Spira for project users properties
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getUsers(user, projectId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getUsersSuccess)
      .withFailureHandler(errorNetwork)
      .getUsers(user, projectId);
  } else {
    msOffice.getUsers(user, projectId)
      .then((response) => getUsersSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets user data on the model
function getUsersSuccess(data) {
  // clear old values
  uiSelection.projectUsers = [];
  // map through user obj and assign names
  uiSelection.projectUsers = data.map(function (item) {
    return {
      id: item.UserId,
      username: item.UserName,
      name: item.FullName,
    };
  });
  model.projectGetRequestsMade++;
}



// check to see that all project and artifact requests have been made - ie that successes match required requests
// returns boolean
function allGetsSucceeded() {
  var projectGetsDone = model.projectGetRequestsToMake === model.projectGetRequestsMade,
    artifactGetsDone = model.artifactGetRequestsToMake === model.artifactGetRequestsMade;
  return projectGetsDone && artifactGetsDone;
}



// send data to server to manage the creation of the template on the relevant sheet
function templateLoader() {
  // set the model based on data stored based on current dropdown selections
  model.currentProject = uiSelection.currentProject;
  model.currentArtifact = uiSelection.currentArtifact;
  model.currentOperation = uiSelection.currentOperation;
  model.currentTemplate = uiSelection.currentTemplate;
  model.isAdmin = isAdmin;
  
  model.projectComponents = [];
  model.projectActiveReleases = [];
  model.projectReleases = [];
  model.projectUsers = [];
  model.projectComponents = uiSelection.projectComponents;
  model.projectActiveReleases = uiSelection.projectActiveReleases;
  model.projectReleases = uiSelection.projectReleases;
  model.projectUsers = uiSelection.projectUsers;
  // get variables ready
  var customs = uiSelection.artifactCustomFields,
    fields = templateFields[model.currentArtifact.field],
    hasBespoke = fieldsWithBespokeData(fields);

  // add bespoke data to relevant fields 
  if (hasBespoke) {
    fields.filter(function (a) {
      var bespokeFieldHasValues = typeof uiSelection[model.currentArtifact.field][a.field] != "undefined" &&
        uiSelection[model.currentArtifact.field][a.field].values;
      return bespokeFieldHasValues;
    }).map(function (field) {
      if (field.bespoke) {
        field.values = uiSelection[model.currentArtifact.field][field.field].values;
      }
      return field;
    });
  }

  if (isGoogle) {
    if (!advancedMode) {
      //if not in advanced mode, ignore the fields only available for that mode
      fields = fields.filter(function (item) {
        if (!item.isAdvanced) {
          return item;
        }
      })
    }
  }

  // collate standard fields and custom fields
  model.fields = fields.concat(customs);

  // get rid of any dropdowns that don't have any values attached
  model.fields = model.fields.filter(function (field) {
    var isNotDrop = field.type !== params.fieldType.drop;
    return isNotDrop || field.values.length > 0;
  });

  // call server side template function
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(templateLoaderSuccess)
      .withFailureHandler(errorImpExp)
      .templateLoader(model, params.fieldType);
  } else {
    msOffice.templateLoader(model, params.fieldType, advancedMode)
      .then(response => templateLoaderSuccess(response))
      .catch(error => error.description ? errorExcel(error) : errorNetwork(error));
  }
}



// once template is loaded, enable the "send to Spira" button
function templateLoaderSuccess(response) {
  model.isTemplateLoaded = true;

  //turn off ajax spinner if it's on
  hideLoadingSpinner();

  // if we get a response string back from server then that means the template was not fully loaded 
  if (response && response.isTemplateLoadFail) {
    return;
  }

  // if we are trying to get data from Spira (ie we clicked the button do so that kicked off loading the template before getting the data itself, get it now)
  if (model.isGettingDataAttempt) {
    getFromSpiraAttempt();
  }
  //enable the send to spira button
  document.getElementById("btn-toSpira").innerHTML = "Add To Spira";
  document.getElementById("btn-toSpira").title = "Create entered data in SpiraPlan"

  //show text in the sidebar that tells the user what the template is set to:
  document.getElementById("template-project").textContent = model.currentProject.name;
  document.getElementById("template-artifact").textContent = model.currentArtifact.name;

}









/*
* 
* ==============
* ERROR HANDLING
* ==============
*
* These call a popup using google server side code
* most args are the HTTPResponse objects from the `withFailureHandler` promise
*
*/
function errorPopUp(type, err) {
  if (isGoogle) {
    google.script.run.error(type);
    //sets the UI to correspond to this mode
    artifactUpdateUI(UI_MODE.errorMode);
    hideLoadingSpinner();
  } else {
    msOffice.error(type, err);
    //sets the UI to correspond to this mode
    artifactUpdateUI(UI_MODE.errorMode);

    if (err != null) {
      console.error("SpiraPlan Import/Export Plugin encountered an error:", err.status ? err.status : "", err.response ? err.response.text : "", err.description ? err.description : "")
    }
    console.info("SpiraPlan Import/Export Plugin: full error is... ", err)
  }
  hideLoadingSpinner();
}

function errorNetwork(err) {
  hideLoadingSpinner();
  errorPopUp("network", err);
}
function errorImpExp(err) {
  errorPopUp('impExp', err);
}
function errorUnknown(err) {
  errorPopUp('unknown', err);
}
function errorExcel(err) {
  errorPopUp('excel', err);
}