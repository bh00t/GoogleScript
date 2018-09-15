var ui = SpreadsheetApp.getUi(); // get UI
var DS_sheet = SpreadsheetApp.getActive().getSheetByName("Team DS"); // get Sheet by Nname

var C_MAX_RESULTS = 1431;

function restCall()
{
   
  var loginDetails = PropertiesService.getUserProperties().getProperty("LoginDetail");
  
  if (loginDetails == null || loginDetails == '')
  {
    ui.alert('Please configure JIRA');
    return ;
  }
  else
  {
    var options = { "Accept":"application/json", 
                   "Content-Type":"application/xml", 
                   "method": "GET",
                   "headers": {"Authorization": "Basic cmlzaGFiaC50cmlwYXRoaTpiaG9vdGFkZXB0aWE="},
                   "muteHttpExceptions": true,
                  // "payload":data
                  };

    var teamLead = DS_sheet.getRange("C1").getValue().toUpperCase();

    switch (teamLead) 
    {
      case "ANSHU":
          teamLead = "ajaiswal";
          break; 
      case "KUSHAGRA":
          teamLead = "kushagra.sharma%40adeptia.com";
          break; 
      case "PIYUSH":
          teamLead = "piyush.gaur";
          break;
      default: 
          teamLead = "rishabh.tripathi";
    }

    
    var url = "jira.adeptia.com/rest/api/2/search?jql=project%20%3D%20TDP%20AND%20issuetype%20%3D%20%22DS%20Issue%20Type%22%20AND%20%22DS%20Sub%20Lead%22%20in%20(%22"+teamLead+"%22)"+"&maxResults="+C_MAX_RESULTS;
    //var url = "jira.adeptia.com/rest/api/2/issue/TDP-19970";//TDP-19970
    var result = UrlFetchApp.fetch(url, options);
    var responseCode = result.getResponseCode();

    //ui.alert(result.getResponseCode() + " blabla " + loginDetails);

   if(responseCode == "200")
   {

      var params = JSON.parse(result.getContentText());    

      var output = JSON.stringify(params);

      var issuesNo = params.issues.length;
      //ui.alert(issuesNo+" blabla " + teamLead);
      var DS_details = new Array(issuesNo);
       DS_sheet.getRange("A3:S1000").clearContent();
      for (i=0; i < issuesNo; i++)
      { 
         DS_details[i] = new Array(19);

         //Jira ID
         var TDP_ID = params.issues[i].key;
         DS_details[i][0] = "=HYPERLINK(\"http://jira.adeptia.com/browse/"+TDP_ID+"\",\""+TDP_ID+"\")";

         //DS Name
         DS_details[i][1] = params.issues[i].fields.customfield_11106;
         //Onboarding Owner
         var onboardingOwner = params.issues[i].fields.customfield_11100;
         DS_details[i][2] = onboardingOwner == null ? "" : onboardingOwner.hasOwnProperty("displayName") == true ? onboardingOwner.displayName : "";
         //WI ID
         DS_details[i][3] = params.issues[i].fields.customfield_11165 == null ? "" : params.issues[i].fields.customfield_11165.toUpperCase().indexOf("ENTER") == -1 ? params.issues[i].fields.customfield_11165 : "";
         //DS Stage
         var DS_stage = params.issues[i].fields.customfield_11141;
         var environment = DS_stage == null ? "" : DS_stage.hasOwnProperty("value") == true ?  DS_stage.value : "";
         var childStage = DS_stage == null ? "" : DS_stage.hasOwnProperty("child") == true ? DS_stage.child.value : "";
         DS_details[i][4] = (environment +" - "+childStage ).trim();
         //Doubtsheet Link
         DS_details[i][5] = params.issues[i].fields.customfield_11168;
         //Specialist
         DS_details[i][6] = params.issues[i].fields.customfield_11171;
         //Site 
         var site = params.issues[i].fields.customfield_11151;
         DS_details[i][7] = site == null ? "" : site.hasOwnProperty("value") == true ? site.value : "";
         //TargetSystem
         var targetSystem = JSON.stringify(params.issues[i].fields.customfield_11149);
         targetSystem = targetSystem.split(":\"")[2];
        targetSystem = targetSystem == null ? "" : targetSystem;
         DS_details[i][8] = targetSystem.substr(0,targetSystem.indexOf("\""));//params.issues[i].fields.customfield_11149.hasOwnProperty("value") == true ? params.issues[i].fields.customfield_11149.value : "";
         //PartnerProfileID
         DS_details[i][9] = params.issues[i].fields.customfield_11161 == null ? "" : params.issues[i].fields.customfield_11161.toUpperCase().indexOf("ENTER") == -1 ? params.issues[i].fields.customfield_11161 : "";
         //start Time
         var startTime = params.issues[i].fields.customfield_11142+"";
        startTime = startTime == null ? "": startTime;
         DS_details[i][10] = startTime.substr(0,startTime.indexOf("T"));
         //end time
         var endTime = params.issues[i].fields.customfield_11143+"";
        endTime = endTime == null ? "" : endTime;
         DS_details[i][11] = endTime.substr(0,endTime.indexOf("T"));
         //Sprint
         var sprint = params.issues[i].fields.customfield_10104;
         sprint = sprint == null ? "": sprint[0]; 
         sprint = sprint.substr(sprint.indexOf(",name="));
         sprint = sprint.substr(sprint.indexOf("="));
         DS_details[i][12] = sprint.substr(1,sprint.indexOf("(")-1);
         //Planned Start Time
         var plannedStartTime = params.issues[i].fields.customfield_11174+"";
         DS_details[i][13] = plannedStartTime;//.substr(0,plannedStartTime.indexOf("T"));
         //Planned End Time
         var plannedEndTime = params.issues[i].fields.customfield_11175+"";
         DS_details[i][14] = plannedEndTime;//.substr(0,plannedEndTime.indexOf("T"));
         //Complexity
         var complexity = params.issues[i].fields.customfield_11120
         DS_details[i][15] = complexity == null ? "": complexity.hasOwnProperty("value") == true ? complexity.value : "";
         //Self Link
         DS_details[i][16] = params.issues[i].self;
         //QA Subtask Link
         DS_details[i][17] = params.issues[i].fields.subtasks.length > 10 ? params.issues[i].fields.subtasks[10].self : ""
         //UAT Subtask Link
         DS_details[i][18] = params.issues[i].fields.subtasks.length > 12 ? params.issues[i].fields.subtasks[12].self : "" ;
      }
      DS_sheet.getRange("A3:S"+(3+issuesNo-1)).setValues(DS_details);

      return;
    }
  }
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntries = [
    {name: "Configure Jira", functionName: "jiraConfigure"},
    {name: "Refresh Sheet", functionName: "restCall"},
    {name: "Unconfigure JIRA", functionName: "removeJiraConfiguration"}
  ]; 
  ss.addMenu("Jira", menuEntries);
  
}

function jiraConfigure()
{
  var userAndPassword = Browser.inputBox("Enter your Jira On Demand User id and Password in the form User:Password. e.g. Tommy.Smith:ilovejira (Note: This will be base64 Encoded and saved as a property on the spreadsheet)", "Userid:Password", Browser.Buttons.OK_CANCEL);
  var x = Utilities.base64Encode(userAndPassword);
  PropertiesService.getUserProperties().setProperty("LoginDetail", "Basic " + x);
}

function removeJiraConfiguration()
{
  var loginDetails = PropertiesService.getUserProperties().getProperty("LoginDetail");
  
  if (loginDetails == null || loginDetails == '')
  {
    ui.alert('JIRA is not configured yet');
    return;
  }
  else
  {
    var result = PropertiesService.getUserProperties().deleteProperty("LoginDetail");
    ui.alert('JIRA is configuration removed' + ' '+result);
    return;
  }
}
