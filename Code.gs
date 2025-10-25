const TASKS_PROPERTY = 'EISENHOWER_TASKS_JSON';
const ID_COUNTER_PROPERTY = 'EISENHOWER_ID_COUNTER';
const GEMINI_KEY_PROPERTY = 'EISENHOWER_GEMINI_API_KEY';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Eisenhower Matrix')
    .addItem('Open Task Manager', 'showTaskManager')
    .addToUi();
}

function showTaskManager() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Eisenhower Matrix')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getTasks() {
  const props = PropertiesService.getDocumentProperties();
  const tasksJson = props.getProperty(TASKS_PROPERTY);
  const idCounterValue = Number(props.getProperty(ID_COUNTER_PROPERTY));

  let tasks = [];
  if (tasksJson) {
    try {
      const parsed = JSON.parse(tasksJson);
      if (Array.isArray(parsed)) {
        tasks = parsed;
      }
    } catch (error) {
      console.warn('Unable to parse stored tasks JSON:', error);
      tasks = [];
    }
  }

  const idCounter = Number.isFinite(idCounterValue) && idCounterValue > 0 ? idCounterValue : 1;

  return {
    tasks,
    idCounter,
  };
}

function saveTasks(payload) {
  const props = PropertiesService.getDocumentProperties();
  const tasks = payload && Array.isArray(payload.tasks) ? payload.tasks : [];
  const idCounter = payload && payload.idCounter ? Number(payload.idCounter) : 1;

  props.setProperty(TASKS_PROPERTY, JSON.stringify(tasks));
  props.setProperty(ID_COUNTER_PROPERTY, String(Number.isFinite(idCounter) && idCounter > 0 ? idCounter : 1));

  return true;
}

function getGeminiKey() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty(GEMINI_KEY_PROPERTY) || '';
}

function setGeminiKey(key) {
  const props = PropertiesService.getDocumentProperties();
  if (key) {
    props.setProperty(GEMINI_KEY_PROPERTY, key);
  } else {
    props.deleteProperty(GEMINI_KEY_PROPERTY);
  }
  return true;
}
