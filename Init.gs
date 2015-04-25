function init_(sDate, eDate) {
    var dayOfEmailsSent = [];
    var dayOfEmailsReceived = [];
    var timeOfEmailsSent = [];
    var timeOfEmailsReceived = [];
    for (i = 0; i < 24; i++) {
        timeOfEmailsSent[i] = 0;
        timeOfEmailsReceived[i] = 0;
    }
    var dayOfWeek = [];
    for (i = 0; i < 14; i++) {
        dayOfWeek[i] = 0;
    }
    var nbrOfEmailsPerConversation = [];
    for (i = 0; i < 101; i++) {
        nbrOfEmailsPerConversation[i] = 0;
    }
    var timeBeforeFirstResponse = [];
    for (i = 0; i < 12; i++) {
        timeBeforeFirstResponse[i] = 0;
    }
    var messagesLength = [];
    for (i = 0; i < 12; i++) {
        messagesLength[i] = 0;
    }
    var topThreads = [];
    for (i = 0; i < 10; i++) {
        topThreads.push(['', 0]);
    }

    var userTimeZone = CalendarApp.getDefaultCalendar().getTimeZone();
    var user = Session.getEffectiveUser().getEmail();
    var variables = {
        range: 0,
        nbrOfConversations: 0,
        nbrOfEmailsPerConversation: nbrOfEmailsPerConversation,
        nbrOfConversationsYouveRepliedTo: 0,
        nbrOfConversationsStartedByYou: 0,
        nbrOfConversationsStarred: 0,
        nbrOfConversationsMarkedAsImportant: 0,
        nbrOfConversationsInInbox: 0,
        nbrOfConversationsInLabels: 0,
        nbrOfConversationsArchived: 0,
        nbrOfConversationsInTrash: 0,
        nbrOfEmailsReceived: 0,
        nbrOfEmailsSent: 0,
        sentDirectlyToYou: 0,
        companyname: user.match(/@([^.]+)/)[1],
        sharedInternally: 0,
        nbrOfAttachmentsReceived: 0,
        nbrOfAttachmentsSent: 0,
        attachmentsSharedInternally: 0,
        attachmentsSent: {
          images:0,
          videos:0,
          audios:0,
          texts:0
        },
        attachmentsReceived: {
          images:0,
          videos:0,
          audios:0,
          texts:0
        },
        timeOfEmailsSent: timeOfEmailsSent,
        timeOfEmailsReceived: timeOfEmailsReceived,
        dayOfWeek: dayOfWeek,
        timeBeforeFirstResponse: timeBeforeFirstResponse,
        messagesLength: messagesLength,
        topThreads: topThreads,
        userTimeZone: userTimeZone,
        user: user
    };
    if (sDate != undefined) {
        variables['type'] = 'custom';
        variables['startDate'] = sDate;
        variables['endDate'] = eDate;
        for (i = 0; i < 31; i++) {
            dayOfEmailsSent[i] = 0;
            dayOfEmailsReceived[i] = 0;
        }
        var monthOfEmailsSent = [];
        var monthOfEmailsReceived = [];
        for (i = 0; i < 12; i++) {
            monthOfEmailsSent[i] = 0;
            monthOfEmailsReceived[i] = 0;
        }
        variables['monthOfEmailsSent'] = monthOfEmailsSent;
        variables['monthOfEmailsReceived'] = monthOfEmailsReceived;
        var status = {
            customReport: true,
            reportSent: "no"
        };
    }
    else {
        // Find previous month...
        variables['previous'] = new Date(new Date().getFullYear(), new Date().getMonth() - 1, 2);
        variables['previousMonth'] = variables['previous'].getMonth();
        variables['year'] = variables['previous'].getYear();
        variables['lastDay'] = daysInMonth_(variables['previousMonth'], variables['year']);
        for (i = 0; i < variables['lastDay']; i++) {
            dayOfEmailsSent[i] = 0;
            dayOfEmailsReceived[i] = 0;
        }
        var status = {
            customReport: false,
            previousMonth: variables['previousMonth'],
            reportSent: "no"
        };
    }
    variables['dayOfEmailsSent'] = dayOfEmailsSent;
    variables['dayOfEmailsReceived'] = dayOfEmailsReceived;
    ScriptProperties.setProperty("variables", Utilities.jsonStringify(variables));
    Utilities.sleep(500);
    ScriptProperties.setProperty("status", Utilities.jsonStringify(status));
    var sheets = ss.getSheets();
    sheets[0].clear();
    var lastColumn = sheets[0].getMaxColumns();
    if(lastColumn > 3){
      sheets[0].deleteColumns(3, lastColumn-3);
    }
    sheets[0].getRange("A1:C1").setValues([
        ["People", "Number of emails received by them", "Number of emails sent to them"]
    ]);
    if (sheets[1] == undefined) {
        ss.insertSheet();
    }
    SpreadsheetApp.flush();
    sheets = ss.getSheets();
    lastColumn = sheets[1].getMaxColumns();
    if(lastColumn > 8){
      sheets[1].deleteColumns(8, lastColumn-8);
    }
    sheets[1].clear().getRange("A1:H1").setValues([["Date", "Day of week", "Thread Id", "From", "To", "Cc", "Words count", "Location"]]);
    sheets[1].getRange("A2:A").setNumberFormat("MMM YY");
    SpreadsheetApp.flush();
    var currentTriggers = ScriptApp.getScriptTriggers();
    for(i in currentTriggers){
      ScriptApp.deleteTrigger(currentTriggers[i]);
    }
    ScriptApp.newTrigger('activityReport').timeBased().everyMinutes(5).create();
}