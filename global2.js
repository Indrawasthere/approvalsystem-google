/* ===============================================
   PART 1 — CONSTANTS, MENU, RESET ROW, UTILITIES
   =============================================== */

// Document types (UPDATED)
var DOCUMENT_TYPES = ["ICC", "Commercial Proposal", "MCSA"];

// If you previously had folder mapping, keep them.
// Only update the keys to match new doc types.
var DOCUMENT_TYPE_FOLDERS = {
  "ICC": "1JFJPfirJuCvZuSEKe6KwRXuupRI0hyyb",
  "Commercial Proposal": "1QVTM_oTSQow9N0e1jNIAc-ADK0qvTwtz",
  "MCSA": "1QVTM_oTSQow9N0e1jNIAc-ADK0qvTwtz"
};

// Your web app URL (unchanged)
var WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzCmtfQLp0tm0kiuz7l7cuZ2IXpICReEIvHuTwRCrnlKxQ5DmtDW6-EX0JD1pwGOmbB/exec";

// ===============================
// Spreadsheet Menu
// ===============================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Atreus Approval")
    .addItem("Send Approval", "sendMultiLayerApproval")
    .addItem("Resubmit After Revision", "triggerResubmit")
    .addSeparator()
    .addItem("Reset Row", "resetRow")
    .addSeparator()
    .addItem("Refresh Menu", "onOpen")
    .addToUi();
}

// ===============================
// Reset Row (Original function logic preserved)
// ===============================
function resetRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert("Select a valid data row.");
    return;
  }

  // Columns reset based on original structure
  sheet.getRange(row, 7).setValue("").setNote("").setBackground(null);   // G Notes
  sheet.getRange(row, 8).setValue("").setNote("").setBackground(null);   // H Level 1
  sheet.getRange(row, 9).setValue("").setNote("").setBackground(null);   // I Level 2
  sheet.getRange(row, 10).setValue("").setNote("").setBackground(null);  // J Level 3
  sheet.getRange(row, 14).setValue("").setBackground(null);              // N Current Editor
  sheet.getRange(row, 15).setValue("").setBackground(null);              // O Overall Status

  SpreadsheetApp.getUi().alert("Row has been reset.");
}

// ===============================
// Utility Functions
// ===============================
function getGMT7Time() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
}

function encodeHTML(str) {
  if (!str) return "";
  return String(str).replace(/[&<>"']/g, function (c) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    }[c];
  });
}

function formatFileSize(bytes) {
  if (!bytes) return "0 B";
  var k = 1024;
  var sizes = ["B", "KB", "MB", "GB", "TB"];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return (bytes / Math.pow(k, i)).toFixed(2) + " " + sizes[i];
}

/* =========================================================
   PART 2 — CORE ENGINE (VALIDATION, NEXT APPROVAL, LINK GEN)
   ========================================================= */

// Extract folder ID from link
function extractFolderIdFromUrl(url) {
  if (!url) return null;

  var patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /\/drive\/folders\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match) return match[1];
  }

  return null;
}

// Validate shared folder
function validateGoogleDriveAttachmentWithType(attachmentUrl, documentType) {
  if (!attachmentUrl || attachmentUrl.trim() === "") {
    return { valid: false, message: "No attachment found" };
  }

  var folderId = extractFolderIdFromUrl(attachmentUrl);
  if (!folderId) {
    return { valid: false, message: "Invalid Drive URL" };
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    var fileList = [];
    var count = 0;
    var totalSize = 0;

    while (files.hasNext() && count < 30) {
      var file = files.next();
      var size = file.getSize();
      fileList.push({
        name: file.getName(),
        size: size,
        sizeFormatted: formatFileSize(size),
        url: file.getUrl()
      });

      totalSize += size;
      count++;
    }

    return {
      valid: true,
      fileCount: count,
      fileList: fileList,
      totalSizeFormatted: formatFileSize(totalSize),
      folderName: folder.getName(),
      url: attachmentUrl
    };
  } catch (e) {
    return { valid: false, message: "Access denied or folder not found" };
  }
}

// =========================================================
// DETERMINE NEXT APPROVAL LAYER
// =========================================================

function getNextApprovalLayerAndEmail(
  levelOne,
  levelTwo,
  levelThree,
  levelOneEmail,
  levelTwoEmail,
  levelThreeEmail,
  currentEditor
) {

  // EDIT MODE flows (unchanged except L3 -> Level1)
  if (currentEditor === "REQUESTER" && levelOne === "REJECTED") {
    return { layer: "LEVEL_ONE", email: levelOneEmail, isResubmit: true };
  }

  if (currentEditor === "LEVEL_ONE" && levelTwo === "REJECTED") {
    return { layer: "LEVEL_TWO", email: levelTwoEmail, isResubmit: true };
  }

  // ✅ NEW REQUEST: L3 reject goes to Level 1
  if (currentEditor === "LEVEL_ONE" && levelThree === "REJECTED") {
    return { layer: "LEVEL_THREE", email: levelThreeEmail, isResubmit: true };
  }

  // Normal progression
  if (!levelOne || levelOne === "" || levelOne === "PENDING") {
    return { layer: "LEVEL_ONE", email: levelOneEmail };
  }

  if (levelOne === "APPROVED" && (!levelTwo || levelTwo === "" || levelTwo === "PENDING")) {
    return { layer: "LEVEL_TWO", email: levelTwoEmail };
  }

  if (levelTwo === "APPROVED" &&
     (!levelThree || levelThree === "" || levelThree === "PENDING")) {
    return { layer: "LEVEL_THREE", email: levelThreeEmail };
  }

  if (levelOne === "APPROVED" &&
      levelTwo === "APPROVED" &&
      levelThree === "APPROVED") {
    return { layer: "COMPLETED", email: "" };
  }

  // fallback
  return { layer: "LEVEL_ONE", email: levelOneEmail };
}

// =========================================================
// APPROVAL LINK GENERATOR (unchanged except sanitized)
// =========================================================

function generateMultiLayerApprovalLink(name, email, description, documentType, attachment, layer, action) {
  var timestamp = new Date().getTime();
  var params = {
    action: action,
    name: encodeURIComponent(name || ""),
    email: encodeURIComponent(email || ""),
    project: encodeURIComponent(description || ""),
    docType: encodeURIComponent(documentType || ""),
    attachment: encodeURIComponent(attachment || ""),
    layer: layer,
    timestamp: timestamp
  };

  var qs = Object.keys(params)
    .map(function (k) { return k + "=" + params[k]; })
    .join("&");

  return WEB_APP_URL + "?" + qs;
}

/* =========================================================
   ORIGINAL FUNCTION — sendMultiLayerApproval (restored)
   ========================================================= */

function sendMultiLayerApproval() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert("Select a valid data row.");
    return;
  }

  var requesterName = sheet.getRange(row, 1).getValue();   // Column Name → Requester Name
  var requesterEmail = sheet.getRange(row, 2).getValue();
  var description = sheet.getRange(row, 3).getValue();
  var documentType = sheet.getRange(row, 4).getValue();
  var attachment = sheet.getRange(row, 5).getValue();

  var levelOneEmail = sheet.getRange(row, 8).getNote();
  var levelTwoEmail = sheet.getRange(row, 9).getNote();
  var levelThreeEmail = sheet.getRange(row, 10).getNote();

  var levelOneStatus = sheet.getRange(row, 8).getValue();
  var levelTwoStatus = sheet.getRange(row, 9).getValue();
  var levelThreeStatus = sheet.getRange(row, 10).getValue();

  var currentEditor = sheet.getRange(row, 14).getValue();
  var overallStatus = sheet.getRange(row, 15).getValue();

  if (!requesterName || !requesterEmail || !description) {
    SpreadsheetApp.getUi().alert("Requester, email, or description missing.");
    return;
  }

  // Validate Document Type (UPDATED)
  if (DOCUMENT_TYPES.indexOf(documentType) === -1) {
    SpreadsheetApp.getUi().alert("Invalid document type.");
    return;
  }

  // Validate attachment
  var validation = validateGoogleDriveAttachmentWithType(attachment, documentType);
  if (!validation.valid) {
    SpreadsheetApp.getUi().alert("Attachment Error: " + validation.message);
    return;
  }

  // Determine next layer
  var next = getNextApprovalLayerAndEmail(
    levelOneStatus,
    levelTwoStatus,
    levelThreeStatus,
    levelOneEmail,
    levelTwoEmail,
    levelThreeEmail,
    currentEditor
  );

  if (!next.email) {
    SpreadsheetApp.getUi().alert("No approver email detected.");
    return;
  }

  var approvalLink = generateMultiLayerApprovalLink(
    requesterName,
    requesterEmail,
    description,
    documentType,
    attachment,
    next.layer,
    "approve"
  );

  var emailSent = sendMultiLayerEmail(
    next.email,
    requesterName,
    requesterEmail,
    description,
    documentType,
    attachment,
    approvalLink,
    next.layer,
    validation,
    next.isResubmit
  );

  if (!emailSent) {
    SpreadsheetApp.getUi().alert("Failed to send email.");
    return;
  }

  // Update sheet
  sheet.getRange(row, 15).setValue("PROCESSING").setBackground("#FFF2CC");
  sheet.getRange(row, 14).setValue(next.layer);

  SpreadsheetApp.getUi().alert("Approval request sent to " + next.email);
}


/* =========================================================
   PART 3 — EMAIL TEMPLATES (CLEAN, NO STEPPER)
   ========================================================= */

// Convert layer code to readable label
function getLayerDisplayName(layerCode) {
  var map = {
    LEVEL_ONE: "Level One",
    LEVEL_TWO: "Level Two",
    LEVEL_THREE: "Level Three"
  };
  return map[layerCode] || layerCode;
}

// =========================================================
// SEND EMAIL
// =========================================================

function sendMultiLayerEmail(
  recipientEmail,
  requesterName,
  requesterEmail,
  description,
  documentType,
  attachment,
  approvalLink,
  layer,
  validationResult,
  isResubmit
) {
  try {
    if (!recipientEmail || recipientEmail.indexOf("@") === -1) return false;

    var layerDisplay = getLayerDisplayName(layer);

    // Document Type badge (UPDATED)
    var badgeColors = {
      "ICC": "#3B82F6",
      "Commercial Proposal": "#10B981",
      "MCSA": "#F59E0B"
    };

    var docBadge =
      '<span style="display:inline-block;background:' +
      (badgeColors[documentType] || "#6B7280") +
      ';color:white;padding:4px 10px;border-radius:10px;font-size:12px;font-weight:bold;margin-top:6px;">' +
      documentType +
      '</span>';

    // Attachment validation section (folder)
    var attachHtml = "";
    if (attachment) {
      if (validationResult && validationResult.valid) {
        var filesHtml = validationResult.fileList
          .map(function (f) {
            return (
              '<li style="background:#f2f2f7;padding:6px 10px;border-radius:6px;margin-bottom:6px;">' +
              "<strong>" + f.name + "</strong>" +
              "<div style='font-size:12px;color:#777'>" + f.sizeFormatted + "</div>" +
              "</li>"
            );
          })
          .join("");

        attachHtml =
          '<div style="margin-top:16px;background:#eef7ff;padding:14px;border-radius:10px;border-left:4px solid #3B82F6;">' +
          "<strong>Attachment Folder:</strong><br>" +
          "<a href='" + attachment + "' target='_blank'>" + validationResult.folderName + "</a>" +
          "<br><small>Total: " + validationResult.fileCount + " file(s)</small>" +
          "<ul style='list-style:none;padding-left:0;margin-top:10px;'>" +
          filesHtml +
          "</ul></div>";
      } else {
        attachHtml =
          '<div style="margin-top:16px;background:#ffecec;padding:14px;border-radius:10px;border-left:4px solid #DC2626;">' +
          "<strong>Attachment Error:</strong> " +
          (validationResult ? validationResult.message : "Invalid link") +
          "<br><a href='" + attachment + "' target='_blank'>Open Link</a>" +
          "</div>";
      }
    }

    // Reject link
    var rejectLink = generateMultiLayerApprovalLink(
      requesterName,
      requesterEmail,
      description,
      documentType,
      attachment,
      layer,
      "reject"
    );

    // RESUBMIT notice (unchanged)
    var resubmitNotice = isResubmit
      ? '<div style="background:#FFF4E6;border-left:4px solid #F59E0B;padding:12px;border-radius:8px;margin:16px 0;">This request was previously rejected and has been revised.</div>'
      : "";

    // ============================================
    // CLEAN HTML EMAIL — NO STEPPER
    // ============================================
    var html =
      '<div style="font-family:Inter,sans-serif;background:#f7f9fb;padding:0;margin:0;">' +
      '<div style="max-width:640px;margin:0 auto;background:white;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);overflow:hidden;">' +

      '<div style="background:#326BC6;color:white;padding:22px;text-align:center;">' +
      "<h2 style='margin:0;font-weight:600;'>" +
      (isResubmit ? "Revised Request" : "Approval Request") +
      "</h2>" +
      "<p style='opacity:0.9;margin-top:6px;'>Stage: " + layerDisplay + "</p>" +
      "</div>" +

      '<div style="padding:24px;">' +
      resubmitNotice +

      "<h3 style='margin:0 0 6px 0;'>" + description + "</h3>" +
      docBadge +
      "<p style='font-size:13px;color:#666;margin-top:8px;'>" +
      "<strong>Requester:</strong> " + requesterName + "<br>" +
      requesterEmail +
      "</p>" +

      attachHtml +

      '<div style="margin-top:24px;">' +
      '<a href="' + approvalLink + '" ' +
      'style="padding:12px 18px;background:#2563EB;color:white;text-decoration:none;border-radius:8px;font-weight:600;margin-right:6px;">Approve</a>' +

      '<a href="' + rejectLink + '" ' +
      'style="padding:12px 18px;background:#DC2626;color:white;text-decoration:none;border-radius:8px;font-weight:600;">Reject</a>' +
      "</div>" +

      "<p style='font-size:11px;color:#777;margin-top:14px;'>Link expires in 7 days</p>" +
      "</div></div></div>";

    // Plain text fallback
    var plain =
      "Approval Request (" + layerDisplay + ")\n" +
      "Project: " + description + "\n" +
      "Document Type: " + documentType + "\n" +
      "Requester: " + requesterName + "\n" +
      "Approve: " + approvalLink + "\n" +
      "Reject: " + rejectLink + "\n";

    MailApp.sendEmail({
      to: recipientEmail,
      subject: (isResubmit ? "[RESUBMIT] " : "") + description + " — " + layerDisplay,
      htmlBody: html,
      body: plain,
      name: "Approval System"
    });

    return true;
  } catch (e) {
    Logger.log("ERROR SEND EMAIL: " + e.toString());
    return false;
  }
}

// =========================================================
// SEND "SEND BACK" EMAIL — Reject Result
// =========================================================

function sendSendBackNotification(
  recipientEmail,
  requesterName,
  requesterEmail,
  description,
  documentType,
  layer,
  rejectionNote,
  rejectorName
) {
  try {
    var layerDisplay = getLayerDisplayName(layer);

    var html =
      '<div style="font-family:Inter,sans-serif;max-width:600px;margin:auto;background:white;padding:24px;border-radius:12px;">' +
      "<h2 style='color:#DC2626;margin-top:0;'>Revision Required</h2>" +
      "<p><strong>Project:</strong> " + description + "</p>" +
      "<p><strong>Rejected by:</strong> " + rejectorName + "</p>" +
      "<p><strong>Stage:</strong> " + layerDisplay + "</p>" +
      "<div style='background:#FFF7F7;border-left:4px solid #DC2626;padding:12px;border-radius:6px;margin-top:12px;'>" +
      "<strong>Feedback:</strong><br>" + encodeHTML(rejectionNote) +
      "</div>" +
      "<p style='font-size:12px;color:#777;margin-top:20px;'>Please revise the attached documents and resubmit via the spreadsheet.</p>" +
      "</div>";

    MailApp.sendEmail({
      to: recipientEmail,
      subject: "Revision Required — " + description,
      htmlBody: html,
      body: "Your submission was rejected at " + layerDisplay + ".\n\nReason:\n" + rejectionNote
    });

    return true;
  } catch (e) {
    Logger.log("ERROR SEND BACK EMAIL: " + e.toString());
    return false;
  }
}

/* =========================================================
   PART 4 — WEB APP ROUTES + APPROVAL / REJECTION HANDLERS
   ========================================================= */

// Safety wrapper (hindari Google Docs race condition)
function doGet(e) {
  try {
    return routeRequest(e);
  } catch (err) {
    return createErrorPage("System error: " + err.message);
  }
}

function doPost(e) {
  try {
    var params = {};
    if (e.postData && e.postData.contents) {
      var pairs = e.postData.contents.split("&");
      pairs.forEach(function (pair) {
        var p = pair.split("=");
        if (p.length === 2) {
          params[decodeURIComponent(p[0])] =
            decodeURIComponent(p[1].replace(/\+/g, " "));
        }
      });
    }

    if (params.action === "submit_rejection") {
      return handleRejectionSubmission(params);
    }

    return createErrorPage("Invalid POST request");
  } catch (err) {
    return createErrorPage("System error: " + err.message);
  }
}

function routeRequest(e) {
  var p = e.parameter || {};
  var action = p.action;

  if (action === "approve") return handleMultiLayerApproval(p);
  if (action === "reject") return handleMultiLayerRejection(p);
  if (action === "submit_rejection") return handleRejectionSubmission(p);

  return createErrorPage("Invalid request.");
}

/* =========================================================
   APPROVAL HANDLER
   ========================================================= */

function handleMultiLayerApproval(params) {
  try {
    var name = decodeURIComponent(params.name || "");
    var email = decodeURIComponent(params.email || "");
    var project = decodeURIComponent(params.project || "");
    var docType = decodeURIComponent(params.docType || "");
    var attachment = decodeURIComponent(params.attachment || "");
    var layer = params.layer || "";
    var timestamp = parseInt(params.timestamp, 10) || 0;

    if (!layer || !project) {
      return createErrorPage("Missing required data.");
    }

    // 7 day expiry check
    if (new Date().getTime() - timestamp > 7 * 24 * 60 * 60 * 1000) {
      return createErrorPage("Link has expired (7 days).");
    }

    var updated = updateMultiLayerApprovalStatus(
      name,
      email,
      project,
      layer
    );

    if (!updated) {
      return createErrorPage("Approval failed or was already processed.");
    }

    // Move automatically to next layer
    try {
      if (layer === "LEVEL_ONE") sendNextApprovalAfterLevelOne();
      if (layer === "LEVEL_TWO") sendNextApprovalAfterLevelTwo();
    } catch (e) {
      Logger.log("AUTO NEXT ERROR: " + e.toString());
    }

    return createSuccessPage(project, docType, layer);
  } catch (err) {
    return createErrorPage("System error: " + err.message);
  }
}

function updateMultiLayerApprovalStatus(name, email, project, layer) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var last = sheet.getLastRow();
  var data = sheet.getRange("A2:O" + last).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    var rowProject = row[2];
    var overall = row[14];

    if (!rowProject) continue;

    if (
      String(rowProject).trim().toLowerCase() !==
      String(project).trim().toLowerCase()
    )
      continue;

    if (
      overall !== "PROCESSING" &&
      overall !== "ACTIVE" &&
      overall !== "EDITING"
    )
      continue;

    var colMap = { LEVEL_ONE: 8, LEVEL_TWO: 9, LEVEL_THREE: 10 };
    var col = colMap[layer];

    var statusCell = sheet.getRange(i + 2, col);
    if (statusCell.getValue() === "APPROVED") return false;

    // Approve
    statusCell.setValue("APPROVED").setBackground("#90EE90");
    statusCell.setNote(
      "Approved by " +
        getLayerDisplayName(layer) +
        " • " +
        getGMT7Time()
    );

    // Check if fully approved
    var l1 = sheet.getRange(i + 2, 8).getValue();
    var l2 = sheet.getRange(i + 2, 9).getValue();
    var l3 = sheet.getRange(i + 2, 10).getValue();

    if (l1 === "APPROVED" && l2 === "APPROVED" && l3 === "APPROVED") {
      sheet
        .getRange(i + 2, 15)
        .setValue("COMPLETED")
        .setBackground("#90EE90")
        .setNote("All layers approved • " + getGMT7Time());
    } else {
      sheet
        .getRange(i + 2, 15)
        .setValue("PROCESSING")
        .setBackground("#FFF2CC")
        .setNote("Updated • " + getGMT7Time());
    }

    return true;
  }

  return false;
}

/* =========================================================
   REJECTION HANDLER
   ========================================================= */

function handleMultiLayerRejection(params) {
  var name = decodeURIComponent(params.name || "");
  var email = decodeURIComponent(params.email || "");
  var project = decodeURIComponent(params.project || "");
  var docType = decodeURIComponent(params.docType || "");
  var attachment = decodeURIComponent(params.attachment || "");
  var layer = params.layer || "";

  if (!project || !layer) {
    return createErrorPage("Missing rejection data.");
  }

  return createRejectionForm(
    name,
    email,
    project,
    docType,
    attachment,
    layer
  );
}

function handleRejectionSubmission(params) {
  var name = params.name || "";
  var email = params.email || "";
  var project = params.project || "";
  var docType = params.docType || "";
  var attachment = params.attachment || "";
  var layer = params.layer || "";
  var rejectionNote = params.rejectionNote || "";

  if (!rejectionNote.trim()) {
    return createErrorPage("Rejection note is required.");
  }

  var result = updateMultiLayerRejectionStatus(
    project,
    layer,
    rejectionNote
  );

  if (!result.success) {
    return createErrorPage("Rejection failed or already processed.");
  }

  // Send notification email
  sendSendBackNotification(
    result.targetEmail,
    name,
    email,
    project,
    docType,
    layer,
    rejectionNote,
    result.sentBackTo
  );

  return createRejectionSuccessPage(project, rejectionNote);
}

function updateMultiLayerRejectionStatus(project, layer, note) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var last = sheet.getLastRow();
  var data = sheet.getRange("A2:O" + last).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    if (
      String(row[2]).trim().toLowerCase() !==
      String(project).trim().toLowerCase()
    )
      continue;

    var colMap = { LEVEL_ONE: 8, LEVEL_TWO: 9, LEVEL_THREE: 10 };
    var col = colMap[layer];
    var statusCell = sheet.getRange(i + 2, col);

    statusCell
      .setValue("REJECTED")
      .setBackground("#FF6B6B")
      .setNote("Rejected • " + getGMT7Time() + "\n" + note);

    // REJECTION FLOW UPDATE:
    // ✅ L1 → Requester
    // ✅ L2 → Level 1
    // ✅ L3 → Level 1 (UPDATED sesuai permintaan)

    var targetEmail = "";
    var sentBackTo = "";

    if (layer === "LEVEL_ONE") {
      targetEmail = row[1];
      sentBackTo = "Requester";
      sheet.getRange(i + 2, 14).setValue("REQUESTER"); // editor
    }

    if (layer === "LEVEL_TWO") {
      targetEmail = row[10];
      sentBackTo = "Level One";
      sheet.getRange(i + 2, 14).setValue("LEVEL_ONE");
    }

    if (layer === "LEVEL_THREE") {
      targetEmail = row[10]; // ✅ balik ke Level 1
      sentBackTo = "Level One";
      sheet.getRange(i + 2, 14).setValue("LEVEL_ONE");
    }

    sheet
      .getRange(i + 2, 15)
      .setValue("EDITING")
      .setBackground("#FFE0B2");

    return {
      success: true,
      targetEmail: targetEmail,
      sentBackTo: sentBackTo
    };
  }

  return { success: false };
}

/* =========================================================
   HTML PAGES (SUCCESS, ERROR, REJECTION FORM)
   ========================================================= */

function createSuccessPage(project, docType, layer) {
  var html =
    "<div style='font-family:Inter;padding:40px;text-align:center;'>" +
    "<h2 style='color:#10B981;'>✔ Approval Successful</h2>" +
    "<p>Project: <strong>" +
    project +
    "</strong></p>" +
    "<p>Stage: " +
    getLayerDisplayName(layer) +
    "</p>" +
    "<p>You may close this page.</p>" +
    "</div>";

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function createErrorPage(msg) {
  var html =
    "<div style='font-family:Inter;padding:40px;text-align:center;'>" +
    "<h2 style='color:#DC2626;'>✖ Error</h2>" +
    "<p>" +
    encodeHTML(msg) +
    "</p>" +
    "</div>";

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function createRejectionForm(
  name,
  email,
  project,
  docType,
  attachment,
  layer
) {
  var html =
    "<div style='font-family:Inter;padding:30px;max-width:600px;margin:auto;'>" +
    "<h2 style='color:#DC2626;'>Reject & Send Back</h2>" +
    "<p><strong>Project:</strong> " +
    project +
    "</p>" +
    "<p><strong>Stage:</strong> " +
    getLayerDisplayName(layer) +
    "</p>" +
    "<form method='post' action='" +
    WEB_APP_URL +
    "'>" +
    "<input type='hidden' name='action' value='submit_rejection'>" +
    "<input type='hidden' name='name' value='" +
    encodeHTML(name) +
    "'>" +
    "<input type='hidden' name='email' value='" +
    encodeHTML(email) +
    "'>" +
    "<input type='hidden' name='project' value='" +
    encodeHTML(project) +
    "'>" +
    "<input type='hidden' name='docType' value='" +
    encodeHTML(docType) +
    "'>" +
    "<input type='hidden' name='attachment' value='" +
    encodeHTML(attachment) +
    "'>" +
    "<input type='hidden' name='layer' value='" +
    encodeHTML(layer) +
    "'>" +
    "<textarea name='rejectionNote' style='width:100%;height:120px;margin-top:10px;' placeholder='Explain the reason...' required></textarea>" +
    "<br><br>" +
    "<button style='padding:10px 16px;background:#DC2626;color:white;border:none;border-radius:6px;'>Submit</button>" +
    "</form>" +
    "</div>";

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function createRejectionSuccessPage(project, note) {
  var html =
    "<div style='font-family:Inter;padding:40px;text-align:center;'>" +
    "<h2 style='color:#10B981;'>✔ Rejection Submitted</h2>" +
    "<p>Project: <strong>" +
    project +
    "</strong></p>" +
    "<p>Rejection note has been recorded.</p>" +
    "</div>";

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}
