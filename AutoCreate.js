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
    <h3>Autocrat Clone Setup</h3>
    <label>Slides Template URL:</label><br>
    <input type="text" id="template" style="width:100%"><br><br>

    <label>Output Folder URL:</label><br>
    <input type="text" id="folder" style="width:100%"><br><br>

    <label>File Name Field (column header):</label><br>
    <input type="text" id="filename"><br><br>

    <label>Email Field (column header):</label><br>
    <input type="text" id="email"><br><br>

    <label>Email Subject:</label><br>
    <input type="text" id="subject" style="width:100%"><br><br>

    <label>Email Body Template:</label><br>
    <textarea id="body" style="width:100%;height:100px"></textarea><br><br>

    <button onclick="save()">Save</button>

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
  `)
  .setWidth(500)
  .setHeight(600);

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
