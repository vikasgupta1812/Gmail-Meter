function uiApp_() {
    var app = UiApp.createApplication().setTitle('Gmail Meter');
    var mainPanel = app.createVerticalPanel().setId('mainPanel');
    var imageOnLoad = app.createImage('https://lh6.googleusercontent.com/-S87nMBe6KWE/TuB9dR48F0I/AAAAAAAAByQ/0Z96LirzDqg/s27/load.gif');
    imageOnLoad.setVisible(false).setId('imageOnLoad').setStyleAttribute('marginLeft', '220px');
    var grid = app.createGrid(4, 2).setCellSpacing(35).setId('grid');
    // 2 choices: Monthly or Custom report
    // Choice 1:
    var titleChoice1 = app.createLabel('Monthly report');
    applyCSS_(titleChoice1, _Title);
    grid.setWidget(0, 0, titleChoice1);
    var descriptionChoice1 = app.createLabel('Get a new report on the first day of every month, showing statistics for the previous month. The first report will be sent as soon as it is ready.');
    applyCSS_(descriptionChoice1, _Description);
    grid.setWidget(1, 0, descriptionChoice1);
    var buttonChoice1 = app.createButton('Get it', app.createServerHandler('getMonthlyReport_'));
    buttonChoice1.addClickHandler(app.createClientHandler().forEventSource().setEnabled(false).forTargets(imageOnLoad).setVisible(true));
    grid.setWidget(2, 0, buttonChoice1);
    // Choice 2:
    var titleChoice2 = app.createLabel('Custom report');
    applyCSS_(titleChoice2, _Title);
    grid.setWidget(0, 1, titleChoice2);
    var descriptionChoice2 = app.createLabel('Get a report on a specific time range (a week, 2 months, 1 year...)');
    applyCSS_(descriptionChoice2, _Description);
    grid.setWidget(1, 1, descriptionChoice2);
    var buttonChoice2 = app.createButton('Get it', app.createServerHandler('getCustomReport_'));
    buttonChoice2.addClickHandler(app.createClientHandler().forEventSource().setEnabled(false));
    grid.setWidget(2, 1, buttonChoice2);
    var helpLink = app.createAnchor('FAQ', 'https://docs.google.com/document/d/1bqLztG7hkqzFx86CXPyJaSgBxisyGtFYQRGYYSKxtdw/edit');
    grid.setWidget(3, 1, helpLink);
    mainPanel.add(grid).add(imageOnLoad);
    app.add(mainPanel);
    ss.show(app);
}

function getMonthlyReport_() {
    var app = UiApp.getActiveApplication();
    init_();
    getReport_(app);
    return app;
}

function getCustomReport_() {
    var app = UiApp.getActiveApplication();
    var grid = app.getElementById('grid').clear().resize(3, 3).setCellSpacing(20);
    grid.setWidget(0, 0, app.createLabel('Start date:').setStyleAttribute('fontWeight', 'bold'));
    grid.setWidget(1, 0, app.createDateBox().setName('sDate'));
    grid.setWidget(0, 1, app.createLabel('End date:').setStyleAttribute('fontWeight', 'bold'));
    grid.setWidget(1, 1, app.createDateBox().setName('eDate'));
    var errorWarning = app.createLabel().setId('error').setText('Wrong Time Range').setVisible(false);
    errorWarning.setStyleAttribute('fontWeight', 'bold').setStyleAttribute('color', 'red');
    grid.setWidget(2, 0, errorWarning);
    var serverHandler = app.createServerHandler('checkAndStart_').addCallbackElement(grid);
    var clientHandler = app.createClientHandler().forEventSource().setEnabled(false)
    .forTargets(errorWarning).setVisible(false)
    .forTargets(app.getElementById('imageOnLoad')).setVisible(true);
    var button = app.createButton('Start').setId('button').addClickHandler(serverHandler);
    button.addClickHandler(clientHandler);
     grid.setWidget(1, 2, button);
    return app;
}

function checkAndStart_(e) {
    var app = UiApp.getActiveApplication();
    var timeRange = new Date(e.parameter.eDate).getTime() - new Date(e.parameter.sDate).getTime();
    if (timeRange > 0) {
        app.getElementById('button').setText('yes');
        init_(e.parameter.sDate, e.parameter.eDate);
        getReport_(app);
    }
    else {
        app.getElementById('error').setVisible(true);
        app.getElementById('button').setEnabled(true);
        app.getElementById('imageOnLoad').setVisible(false);
    }
    return app;
}

function getReport_(app){
    var mainPanel = app.getElementById('mainPanel').clear();
    var subPanel = app.createVerticalPanel().setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
    var logo = app.createImage('https://lh5.googleusercontent.com/-z3DV43-Qc6M/UEO6c3ly0SI/AAAAAAAAC_4/O2QK8LN2lGI/s532/Gmail+Meter+Logo.png').setWidth(250);
    var comment = app.createLabel('Thank you. Gmail Meter is working on your report.').setStyleAttribute('marginTop', '20px');
    var comment2 = app.createLabel('You will receive an email as soon as it is ready.');
    var comment3 = app.createLabel('(In the meantime, you can close this spreadsheet)').setStyleAttribute('marginBottom', '20px');
    var handler = app.createServerHandler('closeUI_');
    var close = app.createButton().setText('Close').addClickHandler(handler);
    subPanel.setStyleAttribute('marginLeft', '100px').add(logo).add(comment).add(comment2).add(comment3).add(close);
    mainPanel.add(subPanel);
    return app;
}

function closeUI_(){
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

// Idea from James Ferreira
// https://sites.google.com/site/scriptsexamples/
function applyCSS_(element, style) {
    for (var key in style) {
        element.setStyleAttribute(key, style[key]);
    }
}

var _Title = {
    "fontSize": "20px",
    "fontWeight": "bold",
    "paddingLeft": "20px",
    "paddingRight": "20px"
}

var _Description = {
    "color": "grey",
    "width": "180px"
}
