// CONFIGURATION
const SPREADSHEET_ID = 'MODIFY_THIS';
const ROOT_FOLDER_ID = 'MODIFY_THIS';
const BATCH_SIZE = 1000; // Fallback if time tracking fails
const STATE_KEY = 'FOLDER_INDEX_STATE'; 
const TIME_KEY = 'LAST_RUN_TIMESTAMP';

// MODIFIED STATE STRUCTURE WITH TIME TRACKING
function indexDriveFolderStructurePaginated() {
  const props = PropertiesService.getScriptProperties();
  const scriptStartTime = Date.now();
  const TIME_LIMIT = 5.5 * 60 * 1000; // 5.5 minutes in milliseconds

  // Load or initialize state
  let state = JSON.parse(props.getProperty(STATE_KEY)) || {
    queue: [],
    processedFolders: [],
    maxDepthFound: 1,
    lastProcessedTime: scriptStartTime
  };

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheets()[0];
    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);

    // Initialize queue if first run
    if (state.queue.length === 0) {
      initSheet(sheet, state.maxDepthFound);
      state.queue.push({
        folderId: rootFolder.getId(),
        path: [],
        depth: 1
      });
    }

    ensureHeaderColumns(sheet, state.maxDepthFound);

    let processedCount = 0;
    let batchData = [];
    let batchMaxDepth = state.maxDepthFound;

    // MAIN PROCESSING LOOP
    while (state.queue.length > 0 && !isTimeUp(scriptStartTime, TIME_LIMIT)) {
      const current = state.queue.shift();
      const folder = DriveApp.getFolderById(current.folderId);
      const folderName = folder.getName();

      // Update depth tracking
      if (current.depth > batchMaxDepth) batchMaxDepth = current.depth;

      // Process files with time checks
      const files = folder.getFiles();
      while (files.hasNext() && !isTimeUp(scriptStartTime, TIME_LIMIT)) {
        const file = files.next();
        batchData.push(createFileRow(file, current.path, current.depth, folderName));
      }

      // Process subfolders with time checks
      const subFolders = folder.getFolders();
      while (subFolders.hasNext() && !isTimeUp(scriptStartTime, TIME_LIMIT)) {
        const subFolder = subFolders.next();
        const subDepth = current.depth + 1;
        state.queue.push({
          folderId: subFolder.getId(),
          path: [...current.path, folder.getName()],
          depth: subDepth
        });
        if (subDepth > batchMaxDepth) batchMaxDepth = subDepth;
      }

      state.processedFolders.push(current.folderId);
      processedCount++;
      state.lastProcessedTime = Date.now(); // Update timestamp
    }

    // Update headers if depth increased
    if (batchMaxDepth > state.maxDepthFound) {
      state.maxDepthFound = batchMaxDepth;
      ensureHeaderColumns(sheet, state.maxDepthFound);
    }

    // Pad and write data
    const paddedData = batchData.map(row => padRow(row, state.maxDepthFound));
    if (paddedData.length > 0) writeToSheet(sheet, paddedData);

    // Save state or cleanup
    if (state.queue.length > 0) {
      props.setProperty(STATE_KEY, JSON.stringify(state));
      props.setProperty(TIME_KEY, state.lastProcessedTime.toString());
      scheduleNextRun();
    } else {
      cleanupState(props);
    }

  } catch (error) {
    handleError(props, state, error);
  }
}

// NEW TIME CHECK FUNCTION
function isTimeUp(startTime, limit) {
  return Date.now() - startTime > limit;
}

// MODIFIED SCHEDULER WITH TIME AWARE RESUME
function scheduleNextRun() {
  // Delete any existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Schedule new trigger with 1-second delay
  ScriptApp.newTrigger('indexDriveFolderStructurePaginated')
    .timeBased()
    .after(1000)
    .create();
}

// MODIFIED PADDING FUNCTION
function padRow(row, maxDepth) {
  const levels = row.slice(0, -4);
  while (levels.length < maxDepth) levels.push('');
  return levels.concat(row.slice(-4));
}

// UPDATED CLEANUP FUNCTION
function cleanupState(props) {
  props.deleteProperty(STATE_KEY);
  props.deleteProperty(TIME_KEY);
  sortSheetHierarchically(); // Add this line
  console.log('Indexing completed successfully');
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function handleError(props, state, error) {
  console.error('Indexing error:', error);
  props.setProperty(STATE_KEY, JSON.stringify(state));
  throw error;
}

function initSheet(sheet, maxDepth) {
  sheet.clearContents();
  const headers = Array.from({length: maxDepth}, (_, i) => `Level ${i+1}`)
    .concat(['File Name', 'Last Updated', 'Size', 'Link']);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function ensureHeaderColumns(sheet, requiredDepth) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentDepth = headers.filter(h => h.startsWith('Level ')).length;

  if (requiredDepth > currentDepth) {
    sheet.insertColumnsAfter(currentDepth, requiredDepth - currentDepth);
    const newHeaders = Array.from({length: requiredDepth - currentDepth}, 
      (_, i) => `Level ${currentDepth + i + 1}`);
    sheet.getRange(1, currentDepth + 1, 1, newHeaders.length)
      .setValues([newHeaders]);
  }
}

// MODIFIED FILE ROW CREATION
function createFileRow(file, path, currentDepth, folderName) {
  const levels = [];
  for (let i = 0; i < currentDepth - 1; i++) {
    levels.push(path[i] || '');
  }
  levels[currentDepth - 1] = folderName;
  return levels.concat([
    file.getName(),
    file.getLastUpdated(),
    formatFileSize(file.getSize()),
    file.getUrl()
  ]);
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  const units = ['KB', 'MB', 'GB'];
  let size = bytes;
  let unitIndex = -1;
  do {
    size /= 1024;
    unitIndex++;
  } while (size >= 1024 && unitIndex < units.length - 1);
  return `${size.toFixed(2)} ${units[unitIndex]}`;
}

function writeToSheet(sheet, batchData) {
  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, batchData.length, batchData[0].length)
    .setValues(batchData);
}
function showStoredState() {
  const props = PropertiesService.getScriptProperties();
  console.log('STATE:', props.getProperty(STATE_KEY));
  console.log('TIME:', props.getProperty(TIME_KEY));
}

function clearAllStorage() {
  PropertiesService.getScriptProperties().deleteAllProperties();
}

