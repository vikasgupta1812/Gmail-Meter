function onInstall(){
  onOpen();
}

function onOpen() {  
  ss.addMenu("Gmail Meter", [{name: "Get a Report", functionName: "uiApp_"}]);
}

function daysInMonth_(month, year) {
    return 32 - new Date(year, month, 32).getDate();
}

function countSendsPerDaysOfWeek_(variables, date, from) {
    var c = 0;
    if (from == variables.user) {
        c = 7;
    }
    switch (Utilities.formatDate(date, variables.userTimeZone, "EEEE")) {
    case "Monday":
        variables.dayOfWeek[c + 0]++;
        break;
    case "Tuesday":
        variables.dayOfWeek[c + 1]++;
        break;
    case "Wednesday":
        variables.dayOfWeek[c + 2]++;
        break;
    case "Thursday":
        variables.dayOfWeek[c + 3]++;
        break;
    case "Friday":
        variables.dayOfWeek[c + 4]++;
        break;
    case "Saturday":
        variables.dayOfWeek[c + 5]++;
        break;
    case "Sunday":
        variables.dayOfWeek[c + 6]++;
        break;
    }
    return variables;
}

function calcMessagesLength_(variables, body, from) {
    var eom = body.indexOf("<pre>");
    if(eom != -1) body = body.substring(0, eom);
    eom = body.indexOf("<div class=\"gmail_quote");
    if(eom != -1) body = body.substring(0, eom);
    eom = body.indexOf("<blockquote class=\"gmail_quote");
    if(eom != -1){
      body = body.substring(0, eom);
      body = body.substring(0, body.lastIndexOf('<br>'));
    }
    var matches = body.replace(/<[^<|>]+?>|&nbsp;/gi, ' ').match(/\b/g);
    var count = 0;
    if (matches) {
        count = matches.length / 2;
    }
    var c = 0;
    if (from == variables.user) {
        c = 6;
    }
    if (count < 10) {
        variables.messagesLength[c + 0]++;
    }
    else if (count < 30) {
        variables.messagesLength[c + 1]++;
    }
    else if (count < 50) {
        variables.messagesLength[c + 2]++;
    }
    else if (count < 100) {
        variables.messagesLength[c + 3]++;
    }
    else if (count < 200) {
        variables.messagesLength[c + 4]++;
    }
    else {
        variables.messagesLength[c + 5]++;
    }
    return [variables, count];
}

function calcWaitingTime_(variables, date, timeOfFirstMessage, youStartedTheConversation) {

    if (youStartedTheConversation) {
        var i = 0;
    }
    else {
        var i = 6;
    }
    var waitingTime = Math.round((date.getTime() - timeOfFirstMessage) / 60000);
    if (waitingTime < 5) {
        variables.timeBeforeFirstResponse[i + 0]++;
    }
    else if (waitingTime < 15) {
        variables.timeBeforeFirstResponse[i + 1]++;
    }
    else if (waitingTime < 60) {
        variables.timeBeforeFirstResponse[i + 2]++;
    }
    else if (waitingTime < 240) {
        variables.timeBeforeFirstResponse[i + 3]++;
    }
    else if (waitingTime < 1440) {
        variables.timeBeforeFirstResponse[i + 4]++;
    }
    else {
        variables.timeBeforeFirstResponse[i + 5]++;
    }
    return variables;
}