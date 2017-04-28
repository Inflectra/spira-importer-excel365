'use strict';

var userInfo = {
  "spiraUrl": null,
  "username": null,
  "apikey": null,
  "auth": null
};

var currentComponents = {

};

var currentUsers = {
  //relevant users will be stored here after selecting a project
};

var currentReleases = {
  //relevant releases will be stored here after selecting a project
};
//template objects for correct order in sheet, currently requirement
//template only

var requirementType = {
  "Package": -1,
  "Need": 1,
  "Feature": 2,
  "Use Case": 3,
  "User Story": 4,
  "Quality": 5,
  "Design Element": 6,
};

var requirementObj = {
  "RequirementId": null,
  "Name": null,
  "Description": null,
  "ReleaseId": null,
  "RequirementTypeId": null,
  "ImportanceId": null,
  "StatusId": null,
  "EstimatePoints": null,
  "AuthorId": null,
  "OwnerId": null,
  "ComponentId": null,
};

var reqStatus = {
  "Requested": 1,
  "Evaluated": 7,
  "Accepted": 5,
  "Rejected": 6,
  "Planned": 2,
  "In Progress": 3,
  "Developed": 4,
  "Obsolete": 8,
  "Tested": 9,
  "Completed": 10,
};

//column ranges for different sheets, currently only requirements.
var columnRanges = {
  requirements: "A3:K",
  customFieldRanges: {
    requirements: ["N", "AQ"],
  }
};

function browserCheck(){
  var isIE = /*@cc_on!@*/false || !!document.documentMode;
  if (isIE){
    var links = $('.breaks-IE');
    for (let i = 0; i < links.length; i++){
      links[i].removeAttribute('target');
    }
    $('#documentation-link').html('For more help, download the documentation by right clicking '
  + '<a href="https://www.inflectra.com/Documents/SpiraTestPlanTeam%20Migration%20and%20Integration%20Guide.pdf" target="_blank">here</a>'
  + ' and selecting "Save Target As...".');
  }
}

//array to hold custom field names
var customFieldNames = [];

function cleanObject(Obj) {
  var cleaned = {};
  for (let i = 0; i < Object.keys(Obj).length; i++) {
    if (Obj[Object.keys(Obj)[i]] != "") {
      cleaned[Object.keys(Obj)[i]] = Obj[Object.keys(Obj)[i]];
    }
  }
  return cleaned;
}

//converts from Excel days since 1/1/1990 to Spira milliseconds since 1/1/1970
function daysToMseconds(days) {
  days -= 2; //for some reason excel returns 2 extra days?
  const between = 25566;
  days -= between;
  let milliseconds = days *= 8.64e+7;
  return '/Date(' + milliseconds + ')/';
}

function disableButtons() {
  $('button').attr('disabled', 'disabled');
  $('button').addClass('is-disabled');
  $('select').attr('disabled', 'disabled');
  $('select').addClass('is-disabled');
}

function enableButtons() {
  $('button').removeClass('is-disabled');
  $('button').prop('disabled', false);
  $('select').removeClass('is-disabled');
  $('select').prop('disabled', false);
}

function multilistConvert(str) {
  function correct(str) {
    let newString = parseInt(str.trim());
    return newString;
  }
  let arr = str.split(",");
  arr = arr.map(correct);
  arr = arr.filter(Number.isInteger);
  return arr;
}

//converts artifact name into a string ending with "Id" to be used
//for accessing the correct keys in objects (ex. turns "requirements" into "RequirementId")
function convertToIdKey(artifactName) {
  let newString = convertToSheetName(artifactName);
  newString = newString.split(" ");
  newString = newString.join("");
  newString = newString.split("");
  newString[newString.length - 1] = "Id";
  newString = newString.join("");
  return newString;
}

//converts artifact name into the correct format for the Excel template sheet names
function convertToSheetName(artifactName) {
  let newString = artifactName.split("-");
  for (let i = 0; i < newString.length; i++) {
    newString[i] = newString[i].charAt(0).toUpperCase()
      + newString[i].substr(1);
  }
  newString = newString.join(" ");
  return newString;
}

(function () {

  function logIn() {
    disableButtons();
    userInfo.spiraUrl = $("#url").val();
    userInfo.username = $("#username").val();
    userInfo.apikey = btoa($("#apikey").val());
    userInfo.auth = btoa("?username=" + $("#username").val() + "&api-key=" + $("#apikey").val());

    //ensures that url has a "/" at the end for doing api calls and adds one if it doesn't
    if (userInfo.spiraUrl.charAt(userInfo.spiraUrl.length - 1) != "/") {
      userInfo.spiraUrl += "/";
    }
    //ensures url is https, otherwise it gives an error and highlights the URL box
    if (userInfo.spiraUrl.charAt(4) !== "s") {
      $('#url').addClass("error");
      $('#error-message').html('<p class="ms-baseFont">Invalid URL (must be https)</p>');
      enableButtons();
    } else {

      getProjects();
    };

  } // end of logIn

  function getProjects() {
    $.ajax({
      method: "GET",
      crossDomain: true,
      url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects' + atob(userInfo.auth),
      success: function (data, textStatus, response) {
        //if call is successful, fill the projects drop down with the user's projects and
        //transition to the main screen
        for (let i = 0; i < data.length; i++) {
          $('<option value="' + data[i].ProjectId + '">' + data[i].Name + '</option>').appendTo('#projects');
        }
        //set current user display, hide log in screen and show main screen.
        $('#current-user').html("Logged in as: " + userInfo.username);
        $('#logInScreen').addClass("hidden");
        $('#mainScreen').removeClass("hidden");
        $('#log-out').removeClass("hidden");
        enableButtons();
      },
      error: function () {
        $('#username').addClass("error");
        $('#apikey').addClass("error");
        $('url').addClass("error");
        $('#error-message').html("Invalid Login Info");
        enableButtons();
      }
    });
  }

  // The initialize function must be run each time a new page is loaded. Currently there is only one page.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      browserCheck();
      $('#logIn').click(logIn);
      $('#clear-log').click(function () {
        $(this).addClass("hidden");
        $('#log-box').html('');
      });

      $('#export').click(function () {
        //When the "send to spira" button is clicked, show the log and loading spinner
        //and then disable the buttons and remove any error coloring that may have
        //been there from a previous attempt to send
        $('#spinner').removeClass("hidden");
        $('#clear-log').removeClass("hidden");
        disableButtons();
        $('#projects').removeClass('error');

        //Check which project they selected and proceed accordingly. -1 is no selected project
        var selectedProject = $('#projects').val();
        if (selectedProject != -1) {
          switch ($('#artifact').val()) {
            case "requirements":
              grabExcelValues(null, $('#artifact').val(), requirementObj, columnRanges.customFieldRanges.requirements);
              break;
            default:
              $('<p class="error-message">Please select an artifact to send.<p>').appendTo('#log-box');
              $('#spinner').addClass("hidden");
              $('#clear-log').removeClass("hidden");
              enableButtons();
          }
        } else {
          $('#projects').addClass('error');
          $('<p class="error-message">No project selected.<p>').appendTo('#log-box');
          enableButtons();
          $('#spinner').addClass("hidden");
        }
      });

      /* The button associated with this function is currently commented out
         in the HTML

      $('#import').click(function () {
        disableButtons();
        switch ($('#artifact').val()) {
          case "requirements":
            ajaxImport($('#artifact').val(), requirementObj);
            break;
          default:
            console.log("Could not import");
        }
      });*/
    })
  };
})();