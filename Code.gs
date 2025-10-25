function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Eisenhower Matrix')
    .addItem('Open Task Manager', 'showTaskManager')
    .addToUi();
}

function showTaskManager() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Eisenhower Matrix')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}
