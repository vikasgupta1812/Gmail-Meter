function sendReport_(variables) {
  var status = Utilities.jsonParse(ScriptProperties.getProperty("status"));
  var nbrOfRecipients = 0;
  var nbrOfSenders = 0;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].sort(2, false);
  var people = sheet.getDataRange().getValues();
  for (var i = 0; i < people.length; i++) {
    if (people[i][1] != 0) nbrOfSenders++;
    if (people[i][2] != 0) nbrOfRecipients++;
  }
  if (nbrOfSenders > 5 && nbrOfRecipients > 5 && variables.nbrOfEmailsSent > 10 && variables.nbrOfEmailsReceived > 10) {
    var report = "<h2 style=\"color:#cccccc; font-family:trebuchet ms;\">Gmail Meter - ";
    if (status.customReport) {
      report += "From " + Utilities.formatDate(new Date(variables.startDate), variables.userTimeZone, 'MM/dd/yyyy');
      report += " to " + Utilities.formatDate(new Date(variables.endDate), variables.userTimeZone, 'MM/dd/yyyy') + "</h2>";
    }
    else {
      report += Utilities.formatDate(new Date(variables.previous), variables.userTimeZone, "MMMM") + "</h2>";
    }
    report += "<br><h3>Volume Statistics</h3>";
    report += "<p style=\"color:#cccccc;\">Chat history, spam, calendar invits and other Google notifications have been removed from this report.</p>";
    report += "<p><table style=\"border-collapse: collapse;\"><tr><td style=\"border: 0px solid white; width: 150px; padding-left: 5px; color: white; background-color: #009754;\"><h3>";
    report += variables.nbrOfConversations + " conversations</h3></td><td style=\"border: 0px solid white; padding-left: 10px;\">";
    report += variables.nbrOfConversationsMarkedAsImportant + " were important.<br>" + variables.nbrOfConversationsStarred + " have been starred.<br>";
    report += "You have started " + Math.round(variables.nbrOfConversationsStartedByYou * 10000 / (variables.nbrOfConversations)) / 100 + "% of them.<br>";
    report += "And have replied to " + Math.round(variables.nbrOfConversationsYouveRepliedTo * 10000 / (variables.nbrOfConversations - variables.nbrOfConversationsStartedByYou)) / 100 + "% of the others.</td></tr></table></p>";


    report += "<p><table style=\"border-collapse: collapse;\"><tr><td style=\"border: 0px solid white; width: 150px; padding-left: 5px; color: white; background-color: #5fb5d8;\"><h3>" + variables.nbrOfEmailsReceived;
    report += " emails received</h3></td><td style=\"border: 0px solid white; padding-left: 10px;\">";
    report += "from " + nbrOfSenders + " people<br>" + Math.round(variables.sentDirectlyToYou * 10000 / variables.nbrOfEmailsReceived) / 100 + "% were sent directly to you.</td></tr></table></p>";

    report += "<p><table style=\"border-collapse: collapse;\"><tr><td style=\"border: 0px solid white; width: 150px; padding-left: 5px; color: white; background-color: #f1481f;\"><h3>" + variables.nbrOfEmailsSent;
    report += " emails sent</h3></td><td style=\"border: 0px solid white; padding-left: 10px;\">";
    report += "to " + nbrOfRecipients + " people<br></td></tr></table></p>";

    report += "<br><h3>Top Senders and Top Recipients</h3>";
    report += "<p style=\"color:#cccccc;\">Mouse over a contact to see the number of emails sent / received.</p>";
    report += "<table style=\"border-collapse: collapse;\"><tr><td style=\"border: 0px solid white;\">" + "<h4>Top 5 senders:</h4><ul>";
    var r = 1;
    var s = 0;
    while (s < 5) {
      if (people[r][0].search(/notification|notify|noreply|update/) == -1) {
        report += "<li title=\"" + people[r][1] + " emails\">" + people[r][0] + "</li>";
        s++;
      }
      r++;
    }
    sheet.sort(3, false);
    var people = sheet.getDataRange().getValues();
    report += "</ul></td><td style=\"border: 0px solid white;\"><h4>Top 5 recipients:</h4><ul>";
    for (i = 1; i < 6; i++) {
      report += "<li title=\"" + people[i][2] + " emails\">" + people[i][0] + "</li>";
    }
    report += "</ul></td></tr></table><br><h3>Daily Traffic</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Time');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Sent');

    var time = '';
    for (var i = 0; i < variables.timeOfEmailsReceived.length; i++) { //create the rows
      switch (i) {
      case 0:
        time = 'Midnight';
        break;
      case 6:
        time = '6 AM';
        break;
      case 12:
        time = 'NOON';
        break;
      case 18:
        time = '6 PM';
        break;
      default:
        time = '';
        break;
      }
      dataTable.addRow([time, variables.timeOfEmailsReceived[i], variables.timeOfEmailsSent[i]]);
    }
    dataTable.build();
    var chartAverageFlow = Charts.newAreaChart().setDataTable(dataTable).setDimensions(650, 400).build();

    report += "cid:Averageflow\'/><h3>Monthly Traffic</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Date');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Sent');

    for (var i = 0; i < variables.dayOfEmailsReceived.length; i++) { //create the rows
      dataTable.addRow([(i + 1).toString(), variables.dayOfEmailsReceived[i], variables.dayOfEmailsSent[i]]);
    }
    dataTable.build();
    var chartDate = Charts.newAreaChart().setDataTable(dataTable).setDimensions(650, 400).build();

    report += "cid:Date\'/><h3>Weekly Traffic</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'dayOfWeek');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Sent');
    var dayTags = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    for (var i = 0; i < dayTags.length; i++) {
      dataTable.addRow([dayTags[i], variables.dayOfWeek[i] * 100 / variables.nbrOfEmailsReceived, variables.dayOfWeek[i + 7] * 100 / variables.nbrOfEmailsSent]);
    }
    dataTable.build();
    var chartDayOfWeek = Charts.newColumnChart().setDataTable(dataTable).setYAxisTitle("in % of messages").setDimensions(650, 400).build();

    report += "cid:DayOfWeek\'/>";

    if (status.customReport) {
      report += "<h3>Month by month</h3><img src=\'";

      var dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType['STRING'], 'Month');
      dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');
      dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Sent');
      var monthTags = ['J', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D'];
      for (var i = 0; i < monthTags.length; i++) {
        dataTable.addRow([monthTags[i], variables.monthOfEmailsReceived[i] * 100 / variables.nbrOfEmailsReceived, variables.monthOfEmailsSent[i] * 100 / variables.nbrOfEmailsSent]);
      }
      dataTable.build();
      var chartMonths = Charts.newColumnChart().setDataTable(dataTable).setYAxisTitle("in % of messages").setDimensions(650, 400).build();

      report += "cid:Months\'/>";
    }

    report += "<h3>Email Categories</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Location');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Number of conversations');
    dataTable.addRow(['In Inbox', variables.nbrOfConversationsInInbox]);
    dataTable.addRow(['In labels', variables.nbrOfConversationsInLabels]);
    dataTable.addRow(['Archived', variables.nbrOfConversationsArchived]);
    dataTable.addRow(['In trash', variables.nbrOfConversationsInTrash]);
    dataTable.build();
    var chartLocation = Charts.newPieChart().setDataTable(dataTable).setDimensions(650, 400).build();

    report += "cid:Location\'/>";

    if (variables.user.search(/gmail|googlemail/) == -1) {

      report += "<h3>For organizations: Internal vs External messages</h3><img src=\'";

      var dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType['STRING'], 'Scope');
      dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Number of emails');
      dataTable.addRow(['Internal discussions', variables.sharedInternally]);
      dataTable.addRow(['Contacts with the outside world', (variables.nbrOfEmailsReceived + variables.nbrOfEmailsSent) - variables.sharedInternally]);
      dataTable.build();
      var chartIntExt = Charts.newPieChart().setDataTable(dataTable).setDimensions(650, 400).build();

      report += "cid:IntExt\'/>";
    }

    report += "<h3>Thread Lengths</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Length');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Number of conversations');

    for (var i = 0; i < variables.nbrOfEmailsPerConversation.length; i++) { //create the rows
      dataTable.addRow([i, variables.nbrOfEmailsPerConversation[i]]);
    }
    dataTable.build();
    var chartConversation = Charts.newScatterChart().setDataTable(dataTable).setDimensions(650, 400).setXAxisLogScale().setYAxisLogScale().setLegendPosition(Charts.Position.NONE).build();

    report += "cid:Conversation\'/><h3>Top threads</h3>";
    report += "<table style=\"border-collapse: collapse; border: 0px solid white;\">";
    for (var i = 0; i < variables.topThreads.length; i++) {
      report += "<tr><td>" + variables.topThreads[i][0] + "</td><td style=\"paddingLeft: 5px;\">" + variables.topThreads[i][1] + "</td></tr>";
    }
    report += "</table><br><h3>Time Before First Response</h3><img src=\'";
    var count = 0;
    for (var k = 0; k < 6; k++) {
      count += variables.timeBeforeFirstResponse[k];
    }
    for (var k = 0; k < 6; k++) {
      if (variables.timeBeforeFirstResponse[k] != 0) variables.timeBeforeFirstResponse[k] = variables.timeBeforeFirstResponse[k] * 100 / count;
    }
    count = 0;
    for (var k = 6; k < 12; k++) {
      count += variables.timeBeforeFirstResponse[k];
    }
    for (var k = 6; k < 12; k++) {
      if (variables.timeBeforeFirstResponse[k] != 0) variables.timeBeforeFirstResponse[k] = variables.timeBeforeFirstResponse[k] * 100 / count;
    }
    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Time');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'When people answer to you');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'When you answer');
    dataTable.addRow(['< 5min', variables.timeBeforeFirstResponse[0], variables.timeBeforeFirstResponse[6]]);
    dataTable.addRow(['< 15min', variables.timeBeforeFirstResponse[1], variables.timeBeforeFirstResponse[7]]);
    dataTable.addRow(['< 1hr', variables.timeBeforeFirstResponse[2], variables.timeBeforeFirstResponse[8]]);
    dataTable.addRow(['< 4hrs', variables.timeBeforeFirstResponse[3], variables.timeBeforeFirstResponse[9]]);
    dataTable.addRow(['< 1day', variables.timeBeforeFirstResponse[4], variables.timeBeforeFirstResponse[10]]);
    dataTable.addRow(['More', variables.timeBeforeFirstResponse[5], variables.timeBeforeFirstResponse[11]]);

    dataTable.build();
    var chartWaitingTime = Charts.newColumnChart().setDataTable(dataTable).setYAxisTitle("in % of messages").setLegendPosition(Charts.Position.BOTTOM).setDimensions(650, 400).build();

    report += "cid:WaitingTime\'/><h3>Word Count</h3><img src=\'";

    count = 0;
    for (var k = 0; k < 6; k++) {
      count += variables.messagesLength[k];
    }
    for (var k = 0; k < 6; k++) {
      if (variables.messagesLength[k] != 0) variables.messagesLength[k] = variables.messagesLength[k] * 100 / count;
    }
    count = 0;
    for (var k = 6; k < 12; k++) {
      count += variables.messagesLength[k];
    }
    for (var k = 6; k < 12; k++) {
      if (variables.messagesLength[k] != 0) variables.messagesLength[k] = variables.messagesLength[k] * 100 / count;
    }

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Words');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'In emails received');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'In emails sent');
    dataTable.addRow(['< 10 words', variables.messagesLength[0], variables.messagesLength[6]]);
    dataTable.addRow(['< 30', variables.messagesLength[1], variables.messagesLength[7]]);
    dataTable.addRow(['< 50', variables.messagesLength[2], variables.messagesLength[8]]);
    dataTable.addRow(['< 100', variables.messagesLength[3], variables.messagesLength[9]]);
    dataTable.addRow(['< 200', variables.messagesLength[4], variables.messagesLength[10]]);
    dataTable.addRow(['More', variables.messagesLength[5], variables.messagesLength[11]]);

    dataTable.build();
    var chartMessagesLength = Charts.newColumnChart().setDataTable(dataTable).setYAxisTitle("in % of messages").setLegendPosition(Charts.Position.BOTTOM).setDimensions(650, 400).build();

    report += "cid:MessagesLength\'/><h3>Attachments</h3>";
    report += "<p>" + variables.nbrOfAttachmentsReceived + " received and " + variables.nbrOfAttachmentsSent + " sent.</p>";

    if (variables.nbrOfAttachmentsReceived > 0) {
      report += "<img src=\'";
      var dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType['STRING'], 'Words');
      dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Attachments received');
      dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Attachments sent');

      for (var contentType in variables.attachmentsSent) {
        if (variables.attachmentsReceived[contentType] > 0 || variables.attachmentsSent[contentType] > 0) {
          dataTable.addRow([contentType, variables.attachmentsReceived[contentType], variables.attachmentsSent[contentType]]);
        }
      }
      dataTable.build();
      var chartAttachmentsTypes = Charts.newColumnChart().setDataTable(dataTable).setYAxisTitle("types of Attachments").setLegendPosition(Charts.Position.BOTTOM).setDimensions(650, 400).build();

      report += "cid:AttachmentsTypes\'/>";
    }

    if (variables.user.search(/gmail|googlemail/) == -1 && variables.nbrOfAttachmentsReceived > 0) {

      report += "<img src=\'";

      var dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType['STRING'], 'Scope');
      dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Number of attachments');
      dataTable.addRow(['Internal sharing of attachments', variables.attachmentsSharedInternally]);
      dataTable.addRow(['Exchanges with the outside world', (variables.nbrOfAttachmentsReceived + variables.nbrOfAttachmentsSent) - variables.attachmentsSharedInternally]);
      dataTable.build();
      var chartIntExtAttachments = Charts.newPieChart().setDataTable(dataTable).setDimensions(650, 400).build();

      report += "cid:IntExtAttachments\'/>";
    }

    report += "<h3>Your Own Statistics</h3>";
    report += "<p>Use <a href=\"" + SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=1\">your data</a> to create your own Charts.</p>";

    var inlineImages = {};
    inlineImages['Averageflow'] = chartAverageFlow;
    inlineImages['Date'] = chartDate;
    inlineImages['DayOfWeek'] = chartDayOfWeek;
    if (status.customReport) {
      inlineImages['Months'] = chartMonths;
    }
    if (variables.user.search(/gmail|googlemail/) == -1) {
      inlineImages['IntExt'] = chartIntExt;
    }
    inlineImages['Location'] = chartLocation;
    inlineImages['Conversation'] = chartConversation;
    inlineImages['WaitingTime'] = chartWaitingTime;
    inlineImages['MessagesLength'] = chartMessagesLength;
    if (variables.nbrOfAttachmentsReceived > 0) {
      inlineImages['AttachmentsTypes'] = chartAttachmentsTypes;
    }
    if (variables.user.search(/gmail|googlemail/) == -1 && variables.nbrOfAttachmentsReceived > 0) {
      inlineImages['IntExtAttachments'] = chartIntExtAttachments;
    }

    MailApp.sendEmail(variables.user, "Gmail Meter", report, {
      htmlBody: report,
      inlineImages: inlineImages
    });
  }
  else {
    var body = "You don't have enough emails to create a report. If you are using aliases,";
    body += "check the FAQ: https://docs.google.com/document/d/1bqLztG7hkqzFx86CXPyJaSgBxisyGtFYQRGYYSKxtdw/edit";
    MailApp.sendEmail(variables.user, "Gmail Meter", body);
  }
  status.reportSent = "yes";
  ScriptProperties.setProperty("status", Utilities.jsonStringify(status));
  var currentTriggers = ScriptApp.getScriptTriggers();
  for (i in currentTriggers) {
    ScriptApp.deleteTrigger(currentTriggers[i]);
  }
  ScriptApp.newTrigger('activityReport').timeBased().everyDays(1).atHour(1).create();
}