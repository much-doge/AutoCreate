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
    .createMenu("Document Engine")
    .addItem("Setup", "showSetupDialog")
    .addItem("Process Rows", "processRows")
    .addToUi();
}

/* ==============================
   SETUP DIALOG
================================ */

function showSetupDialog() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, 
                         "Helvetica Neue", Arial, sans-serif;
            padding: 24px;
            background: #f8f9fa;
            color: #202124;
          }

          h3 {
            margin-top: 0;
            font-weight: 600;
            font-size: 18px;
          }

          label {
            display: block;
            margin-top: 16px;
            margin-bottom: 6px;
            font-size: 13px;
            font-weight: 500;
            color: #5f6368;
          }

          input, textarea {
            width: 100%;
            padding: 8px 10px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            font-size: 13px;
            box-sizing: border-box;
            transition: border 0.2s ease;
          }

          input:focus, textarea:focus {
            outline: none;
            border: 1px solid #1a73e8;
          }

          textarea {
            min-height: 100px;
            resize: vertical;
          }

          button {
            margin-top: 24px;
            width: 100%;
            padding: 10px;
            background: #1a73e8;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: background 0.2s ease;
          }

          button:hover {
            background: #1557b0;
          }

          .container {
            max-width: 480px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h3>Document Engine Setup</h3>

          <label>Slides Template URL</label>
          <input type="text" id="template">

          <label>Output Folder URL</label>
          <input type="text" id="folder">

          <label>File Name Field (column header)</label>
          <input type="text" id="filename">

          <label>Email Field (column header)</label>
          <input type="text" id="email">

          <label>Email Subject</label>
          <input type="text" id="subject">

          <label>Email Body Template</label>
          <textarea id="body"></textarea>

          <button onclick="save()">Save Configuration</button>
        </div>

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
            google.script.run.saveSetup(data);
            google.script.host.close();
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(520)
  .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, "Setup");
}

function saveSetup(data) {
  SCRIPT_PROPERTIES.setProperties(data);
}

/* ==============================
   PROCESSING ENGINE
================================ */

function processRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const templateId = extractId(SCRIPT_PROPERTIES.getProperty("template"));
  const folderId = extractId(SCRIPT_PROPERTIES.getProperty("folder"));
  const filenameField = SCRIPT_PROPERTIES.getProperty("filename");
  const emailField = SCRIPT_PROPERTIES.getProperty("email");
  const subjectTemplate = SCRIPT_PROPERTIES.getProperty("subject");
  const bodyTemplate = SCRIPT_PROPERTIES.getProperty("body");

  const folder = DriveApp.getFolderById(folderId);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowObj = {};

    headers.forEach((h, idx) => {
      rowObj[h] = row[idx];
    });

    const copy = DriveApp.getFileById(templateId).makeCopy(folder);
    const presentation = SlidesApp.openById(copy.getId());

    replacePlaceholders(presentation, rowObj);

    presentation.saveAndClose();

    const pdf = copy.getBlob().getAs("application/pdf");
    const fileName = rowObj[filenameField] + ".pdf";

    const pdfFile = folder.createFile(pdf).setName(fileName);

    pdfFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    const email = rowObj[emailField];

    const subject = replaceText(subjectTemplate, rowObj);
    const body = replaceText(bodyTemplate, rowObj);

    GmailApp.sendEmail(email, subject, body, {
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
  const slides = presentation.getSlides();

  slides.forEach(slide => {
    Object.keys(data).forEach(key => {
      slide.replaceAllText(`<<${key}>>`, data[key]);
    });
  });
}

function replaceText(template, data) {
  let result = template;
  Object.keys(data).forEach(key => {
    result = result.replace(
      new RegExp(`<<${key}>>`, 'g'),
      data[key]
    );
  });
  return result;
}

function extractId(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
