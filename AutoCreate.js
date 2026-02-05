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


const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("AutoCreate")
    .addItem("Setup (This Sheet)", "showSetupDialog")
    .addItem("Process Active Sheet", "processRows")
    .addToUi();
}

/* ==============================
   SETUP DIALOG
================================ */

function showSetupDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const config = getSheetConfig(sheetName);

  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
            padding: 24px;
            background: #f8f9fa;
            color: #202124;
          }
          h3 { margin-top: 0; font-weight: 600; }
          label {
            display: block;
            margin-top: 14px;
            font-size: 13px;
            font-weight: 500;
            color: #5f6368;
          }
          input, textarea {
            width: 100%;
            padding: 8px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            font-size: 13px;
            box-sizing: border-box;
          }
          textarea { min-height: 90px; }
          button {
            margin-top: 20px;
            width: 100%;
            padding: 10px;
            background: #1a73e8;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            cursor: pointer;
          }
        </style>
      </head>
      <body>
        <h3>Setup for Sheet: ${sheetName}</h3>

        <label>Slides Template URL</label>
        <input type="text" id="template" value="${config.template || ""}">

        <label>Output Folder URL</label>
        <input type="text" id="folder" value="${config.folder || ""}">

        <label>File Name Field (column header)</label>
        <input type="text" id="filename" value="${config.filename || ""}">

        <label>Email Field (column header)</label>
        <input type="text" id="email" value="${config.email || ""}">

        <label>Email Subject</label>
        <input type="text" id="subject" value="${config.subject || ""}">

        <label>Email Body Template</label>
        <textarea id="body">${config.body || ""}</textarea>

        <button onclick="save()">Save Configuration</button>

        <script>
          function save() {
            const data = {
              template: document.getElementById("template").value,
              folder: document.getElementById("folder").value,
              filename: document.getElementById("filename").value,
              email: document.getElementById("email").value,
              subject: document.getElementById("subject").value,
              body: document.getElementById("body").value
            };
            google.script.run.saveSetupForSheet(data);
            google.script.host.close();
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(520)
  .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, "Sheet Configuration");
}

function saveSetupForSheet(data) {
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  Object.keys(data).forEach(key => {
    SCRIPT_PROPERTIES.setProperty(`${sheetName}_${key}`, data[key]);
  });
}

function getSheetConfig(sheetName) {
  return {
    template: SCRIPT_PROPERTIES.getProperty(`${sheetName}_template`),
    folder: SCRIPT_PROPERTIES.getProperty(`${sheetName}_folder`),
    filename: SCRIPT_PROPERTIES.getProperty(`${sheetName}_filename`),
    email: SCRIPT_PROPERTIES.getProperty(`${sheetName}_email`),
    subject: SCRIPT_PROPERTIES.getProperty(`${sheetName}_subject`),
    body: SCRIPT_PROPERTIES.getProperty(`${sheetName}_body`)
  };
}

/* ==============================
   PROCESSING ENGINE
================================ */

function processRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const config = getSheetConfig(sheetName);

  if (!config.template || !config.folder) {
    SpreadsheetApp.getUi().alert("This sheet is not configured yet.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const templateId = extractId(config.template);
  const folderId = extractId(config.folder);

  const folder = DriveApp.getFolderById(folderId);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowObj = {};
    headers.forEach((h, idx) => rowObj[h] = row[idx]);

    const copy = DriveApp.getFileById(templateId).makeCopy(folder);
    const presentation = SlidesApp.openById(copy.getId());

    replacePlaceholders(presentation, rowObj);
    presentation.saveAndClose();

    const pdf = copy.getBlob().getAs("application/pdf");
    const fileName = rowObj[config.filename] + ".pdf";
    const pdfFile = folder.createFile(pdf).setName(fileName);

    pdfFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    const subject = replaceText(config.subject, rowObj);
    const body = replaceText(config.body, rowObj);

    GmailApp.sendEmail(rowObj[config.email], subject, body, {
      attachments: [pdfFile]
    });

    sheet.getRange(i + 1, headers.length + 1)
      .setValue(pdfFile.getUrl());

    copy.setTrashed(true);
  }
}

/* ==============================
   HELPERS
================================ */

function replacePlaceholders(presentation, data) {
  presentation.getSlides().forEach(slide => {
    Object.keys(data).forEach(key => {
      slide.replaceAllText(`<<${key}>>`, data[key]);
    });
  });
}

function replaceText(template, data) {
  let result = template || "";
  Object.keys(data).forEach(key => {
    result = result.replace(new RegExp(`<<${key}>>`, 'g'), data[key]);
  });
  return result;
}

function extractId(url) {
  const match = url.match(/[-\\w]{25,}/);
  return match ? match[0] : null;
}
