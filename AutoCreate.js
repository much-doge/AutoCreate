/*
 * AutoCreate (Autocrat-like Google Apps Script)
 * Copyright (C) 2026 Nur Eko Windianto
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */


const CONFIGS_SHEET_NAME = 'AutoCreate_Configs';

/* ========== ON-OPEN: Add menu ========== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("AutoCreate")
    .addItem("Manage Setups", "showConfigManager")
    .addItem("Process a Setup", "chooseAndProcessJob")
    .addToUi();
}

/* ========== CONFIG SHEET ========== */

// Ensure config sheet exists and set up columns
function ensureConfigsSheet() {
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIGS_SHEET_NAME, 0);
    sheet.hideSheet();
    sheet.getRange(1, 1, 1, 8).setValues([[
      "JobName",           // Human-friendly name
      "DataSheet",         // Merge sheet name
      "SlidesTemplateURL", // Slides template URL
      "OutputFolderURL",   // Folder URL for generated files
      "FileNameField",     // Column header for file name
      "EmailField",        // Column header for recipient email
      "EmailSubject",      // Email subject template
      "EmailBody"          // Email body template
    ]]);
  }
  return sheet;
}

/* ========== CONFIG MANAGEMENT UI ========== */

function showConfigManager() {
  // Shows a little help and asks user to unhide then edit config sheet
  ensureConfigsSheet();
  SpreadsheetApp.getUi().alert(
    'To add/edit mail merge jobs, unhide the "' + CONFIGS_SHEET_NAME + '" sheet in the spreadsheet, ' +
    'then add one row for each template/job/config (see column headers). ' +
    '\n\nHide it again after editing if you want.'
  );
}

/* ========== SELECT AND PROCESS A JOB ========== */

function chooseAndProcessJob() {
  const jobs = getAllJobs();
  if (jobs.length === 0) {
    SpreadsheetApp.getUi().alert('No setups found!\n\nUse "Manage Setups" menu to add some first.');
    return;
  }
  // Build selection prompt
  const ui = SpreadsheetApp.getUi();
  let html = '<html><body style="font-family:Arial;padding:1em"><b>Select a setup/job to process:</b><br><br>';
  jobs.forEach((job, idx) => {
    html += `<button style="margin:0.5em 0;padding:6px 16px;" onclick="google.script.run.processJobByIndex(${idx});window.close()">${escapeHtml(job.JobName)}</button><br>`;
  });
  html += '</body></html>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(320).setHeight(100 + jobs.length * 38), 'Run AutoCreate Setup');
}

/* ========== CORE ENGINE: GET JOBS, PROCESS ONE ========== */

// Read all jobs from the config sheet
function getAllJobs() {
  let sheet = ensureConfigsSheet();
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  let jobs = [];
  for (let i = 1; i < data.length; i++) {
    let job = {};
    data[i].forEach((v, j) => job[headers[j]] = v);
    if (job.JobName && job.DataSheet) jobs.push(job);
  }
  return jobs;
}

// Called by HTML menu
function processJobByIndex(idx) {
  let jobs = getAllJobs();
  if (idx < 0 || idx >= jobs.length) throw new Error('Invalid job index');
  processJob(jobs[idx]);
}

// Main processing logic
function processJob(job) {
  // Step 1: Setup and safe extraction
  let templateId = extractId(job.SlidesTemplateURL);
  let folderId   = extractId(job.OutputFolderURL);
  let fileNameField = job.FileNameField;
  let emailField    = job.EmailField;
  let subjectTemplate = job.EmailSubject;
  let bodyTemplate    = job.EmailBody;
  let dataSheet = SpreadsheetApp.getActive().getSheetByName(job.DataSheet);

  if (!dataSheet) throw new Error('Data sheet not found: ' + job.DataSheet);

  // Step 2: Prepare to process data
  let data = dataSheet.getDataRange().getValues();
  let headers = data[0];
  let folder = DriveApp.getFolderById(folderId);

  // Make sure there is a result column for links
  let outputCol = headers.length + 1;
  if (dataSheet.getLastColumn() === headers.length) {
    dataSheet.insertColumnAfter(headers.length);
    dataSheet.getRange(1, outputCol).setValue('Doc/PDF URL');
  }

  let sentCount = 0;

  // Step 3: Process each row (Reasoning: all prep BEFORE sending e-mails)
  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    // Skip blank data rows
    if (!rowData || !rowData.some(val => val && val.toString().trim())) continue;

    // Map data by headers
    let rowObj = {};
    headers.forEach((h, idx) => rowObj[h] = rowData[idx]);

    // Prepare merged doc
    const copy = DriveApp.getFileById(templateId).makeCopy(folder);
    const presentation = SlidesApp.openById(copy.getId());
    replacePlaceholders(presentation, rowObj);
    presentation.saveAndClose();

    // Export as PDF
    const pdf = copy.getBlob().getAs("application/pdf");
    const fileName = (rowObj[fileNameField] || ("Document_" + (i + 1))) + ".pdf";
    const pdfFile = folder.createFile(pdf).setName(fileName);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Compose email
    const emailTo = rowObj[emailField];
    const subject = replaceText(subjectTemplate, rowObj);
    const body    = replaceText(bodyTemplate, rowObj);

    // Record file link in output col
    dataSheet.getRange(i+1, outputCol).setValue(pdfFile.getUrl());

    // SEND (dispatch after all prep)
    GmailApp.sendEmail(emailTo, subject, body, {
      attachments: [pdfFile]
    });

    sentCount++;
    copy.setTrashed(true); // Clean up copied SLIDES
  }

  SpreadsheetApp.getUi().alert('Processed ' + sentCount + ' rows for setup: ' + job.JobName);
}

/* ========== HELPERS ========== */

function replacePlaceholders(presentation, data) {
  const slides = presentation.getSlides();
  slides.forEach(slide => {
    Object.keys(data).forEach(key => {
      slide.replaceAllText(`<<${key}>>`, String(data[key] ?? ''));
    });
  });
}

function replaceText(template, data) {
  let result = template || '';
  Object.keys(data).forEach(key => {
    result = result.replace(new RegExp(`<<${key}>>`, 'g'), data[key]);
  });
  return result;
}

function extractId(url) {
  let match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function escapeHtml(str) {
  return String(str || '').replace(/[&<>"']/g, function (m) {
    return ({'&': '&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[m];
  });
}