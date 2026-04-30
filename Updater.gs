const EXPENSE_TOOLS_LOCAL_VERSION = '1.0.0';

const GITHUB_BASE_URL = 'https://raw.githubusercontent.com/tdickey-stack/expense-tools/main/';
const VERSION_JSON_URL = GITHUB_BASE_URL + 'version.json';

function checkForExpenseToolsUpdates() {
  const ui = SpreadsheetApp.getUi();

  const latest = getLatestExpenseToolsVersion_();
  const currentVersion = getCurrentExpenseToolsVersion_();

  if (compareVersions_(latest.version, currentVersion) <= 0) {
    ui.alert(
      'Expense Tools is up to date.',
      `Current Version: ${currentVersion}`,
      ui.ButtonSet.OK
    );
    return;
  }

  const changelog = latest.changelog
    .map(item => `• ${item}`)
    .join('\n');

  const response = ui.alert(
    'Expense Tools Update Available',
    `Current Version: ${currentVersion}\nLatest Version: ${latest.version}\n\nChangelog:\n${changelog}\n\nUpdate now?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  updateExpenseToolsFromGitHub_();

  ui.alert(
    'Update complete.',
    `Expense Tools has been updated to v${latest.version}.\n\nPlease reload the spreadsheet.`,
    ui.ButtonSet.OK
  );
}

function getCurrentExpenseToolsVersion_() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty('EXPENSE_TOOLS_VERSION') || EXPENSE_TOOLS_LOCAL_VERSION;
}

function setCurrentExpenseToolsVersion_(version) {
  PropertiesService.getDocumentProperties()
    .setProperty('EXPENSE_TOOLS_VERSION', version);
}

function getLatestExpenseToolsVersion_() {
  const response = UrlFetchApp.fetch(VERSION_JSON_URL, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Could not fetch version.json from GitHub.');
  }

  return JSON.parse(response.getContentText());
}

function updateExpenseToolsFromGitHub_() {
  const latest = getLatestExpenseToolsVersion_();

  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();

  const currentProject = getAppsScriptProjectContent_(scriptId, token);
  const currentFiles = currentProject.files || [];

  const updatedFilesByKey = {};

  currentFiles.forEach(file => {
    updatedFilesByKey[file.name + '|' + file.type] = file;
  });

  latest.files.forEach(fileInfo => {
    const sourceUrl = GITHUB_BASE_URL + fileInfo.path;

    const response = UrlFetchApp.fetch(sourceUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Could not fetch ${fileInfo.path} from GitHub.`);
    }

    updatedFilesByKey[fileInfo.name + '|' + fileInfo.type] = {
      name: fileInfo.name,
      type: fileInfo.type,
      source: response.getContentText()
    };
  });

  const updatedFiles = Object.values(updatedFilesByKey);

  const hasManifest = updatedFiles.some(file =>
    file.name === 'appsscript' && file.type === 'JSON'
  );

  if (!hasManifest) {
    throw new Error('Missing appsscript.json manifest file. Update cancelled.');
  }

  updateAppsScriptProjectContent_(scriptId, token, updatedFiles);

  setCurrentExpenseToolsVersion_(latest.version);
}

function getAppsScriptProjectContent_(scriptId, token) {
  const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Could not read Apps Script project content: ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

function updateAppsScriptProjectContent_(scriptId, token, files) {
  const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;

  const response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      files: files
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Could not update Apps Script project: ' + response.getContentText());
  }
}

function compareVersions_(a, b) {
  const aParts = String(a).split('.').map(Number);
  const bParts = String(b).split('.').map(Number);

  for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
    const aPart = aParts[i] || 0;
    const bPart = bParts[i] || 0;

    if (aPart > bPart) return 1;
    if (aPart < bPart) return -1;
  }

  return 0;
}
