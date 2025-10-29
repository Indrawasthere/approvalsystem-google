// ============================================
// PART 1: MAIN FUNCTIONS & CORE LOGIC
// Multi-Layer Reusable Approval System
// ============================================

var DOCUMENT_TYPE_FOLDERS = {
  ICC: "1JFJPfirJuCvZuSEKe6KwRXuupRI0hyyb",
  Quotation: "1QVTM_oTSQow9N0e1jNIAc-ADK0qvTwtz",
  Proposal: "1QVTM_oTSQow9N0e1jNIAc-ADK0qvTwtz",
};

// GANTI INI DENGAN WEB APP URL LU
var WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbwJjiTVQOjoH0WGp85eV8jdtsneOa-sv0vG37XY641497eB5ooNaifKOGaa_lJZXKa1/exec";

// ============================================
// MAIN APPROVAL SENDER
// ============================================

function sendMultiLayerApproval() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("No data to process!");
      return;
    }

    // Column mapping: A-O (15 columns)
    var dataRange = sheet.getRange("A2:O" + lastRow);
    var data = dataRange.getValues();

    var results = [];
    var processedCount = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var name = row[0]; // Column A - Name
      var email = row[1]; // Column B - Requester Email
      var description = row[2]; // Column C - Description
      var documentType = row[3]; // Column D - Document Type
      var attachment = row[4]; // Column E - Attachment
      var sendStatus = row[5]; // Column F - Send Checkbox
      var levelOneStatus = row[7]; // Column H - Level One Status
      var levelTwoStatus = row[8]; // Column I - Level Two Status
      var levelThreeStatus = row[9]; // Column J - Level Three Status
      var levelOneEmail = row[10]; // Column K - Level One Email
      var levelTwoEmail = row[11]; // Column L - Level Two Email
      var levelThreeEmail = row[12]; // Column M - Level Three Email
      var currentEditor = row[13]; // Column N - Current Editor
      var overallStatus = row[14]; // Column O - Overall Status

      // Skip empty rows
      if (!name || !email) continue;

      // Check if send checkbox is checked
      var isChecked =
        sendStatus === true || sendStatus === "TRUE" || sendStatus === "true";

      if (isChecked) {
        Logger.log("Processing: " + name + " - " + description);

        // Validate attachment
        var validationResult = validateGoogleDriveAttachmentWithType(
          attachment,
          documentType
        );

        // Determine next approval layer
        var nextApproval = getNextApprovalLayerAndEmail(
          levelOneStatus,
          levelTwoStatus,
          levelThreeStatus,
          levelOneEmail,
          levelTwoEmail,
          levelThreeEmail,
          currentEditor
        );

        Logger.log(
          "Next approval: " + nextApproval.layer + " -> " + nextApproval.email
        );

        if (nextApproval.layer !== "COMPLETED" && nextApproval.email) {
          var approvalLink = generateMultiLayerApprovalLink(
            name,
            email,
            description,
            documentType,
            attachment,
            nextApproval.layer,
            "approve"
          );

          var emailSent = sendMultiLayerEmail(
            nextApproval.email,
            description,
            documentType,
            attachment,
            approvalLink,
            nextApproval.layer,
            validationResult,
            nextApproval.isResubmit
          );

          if (emailSent) {
            // Update status to PROCESSING
            sheet.getRange(i + 2, 15).setValue("PROCESSING"); // Column O
            sheet.getRange(i + 2, 15).setBackground("#FFF2CC");

            // Log approval link
            sheet
              .getRange(i + 2, 7)
              .setNote(
                "APPROVAL_LINK_" + nextApproval.layer + ": " + approvalLink
              );

            results.push({
              name: name,
              project: description,
              layer: nextApproval.layer,
              status: "Email Sent to: " + nextApproval.email,
              isResubmit: nextApproval.isResubmit,
            });

            processedCount++;
            Logger.log("✅ Approval email sent for: " + name);
          }
        }

        // Reset checkbox after processing
        sheet.getRange(i + 2, 6).setValue(false);
        Utilities.sleep(1000);
      }
    }

    showMultiLayerSummary(results, processedCount);
  } catch (error) {
    Logger.log("Error in sendMultiLayerApproval: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "System Error",
      "Error: " + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================
// CORE LOGIC: DETERMINE NEXT APPROVAL LAYER
// ============================================

function getNextApprovalLayerAndEmail(
  levelOne,
  levelTwo,
  levelThree,
  levelOneEmail,
  levelTwoEmail,
  levelThreeEmail,
  currentEditor
) {
  Logger.log("Checking approval flow:");
  Logger.log(
    "Level One: " +
      levelOne +
      ", Level Two: " +
      levelTwo +
      ", Level Three: " +
      levelThree
  );
  Logger.log("Current Editor: " + currentEditor);

  // Handle editing states first based on Current Editor
  if (currentEditor === "REQUESTER" && levelOne === "REJECTED") {
    Logger.log("Requester editing after Level One rejection");
    return {
      layer: "LEVEL_ONE",
      email: levelOneEmail,
      isResubmit: true,
    };
  }

  if (currentEditor === "LEVEL_ONE" && levelTwo === "REJECTED") {
    Logger.log("Level One editing after Level Two rejection");
    return {
      layer: "LEVEL_TWO",
      email: levelTwoEmail,
      isResubmit: true,
    };
  }

  if (currentEditor === "LEVEL_TWO" && levelThree === "REJECTED") {
    Logger.log("Level Two editing after Level Three rejection");
    return {
      layer: "LEVEL_THREE",
      email: levelThreeEmail,
      isResubmit: true,
    };
  }

  // Normal flow scenarios
  // SCENARIO 1: First submission or pending Level One
  if (!levelOne || levelOne === "" || levelOne === "PENDING") {
    Logger.log("Next: Level One (First submission)");
    return {
      layer: "LEVEL_ONE",
      email: levelOneEmail,
      isResubmit: false,
    };
  }

  // SCENARIO 2: Level One approved, proceed to Level Two
  if (
    levelOne === "APPROVED" &&
    (!levelTwo || levelTwo === "" || levelTwo === "PENDING")
  ) {
    Logger.log("Next: Level Two");
    return {
      layer: "LEVEL_TWO",
      email: levelTwoEmail,
      isResubmit: false,
    };
  }

  // SCENARIO 3: Level Two approved, proceed to Level Three
  if (
    levelTwo === "APPROVED" &&
    (!levelThree || levelThree === "" || levelThree === "PENDING")
  ) {
    Logger.log("Next: Level Three");
    return {
      layer: "LEVEL_THREE",
      email: levelThreeEmail,
      isResubmit: false,
    };
  }

  // SCENARIO 4: All approved - completed
  if (
    levelOne === "APPROVED" &&
    levelTwo === "APPROVED" &&
    levelThree === "APPROVED"
  ) {
    Logger.log("Next: COMPLETED");
    return {
      layer: "COMPLETED",
      email: "",
      isResubmit: false,
    };
  }

  // Default: Stay at current state
  Logger.log("Default: Maintaining current state");
  return {
    layer: "LEVEL_ONE",
    email: levelOneEmail,
    isResubmit: false,
  };
}

// ============================================
// GENERATE APPROVAL LINK
// ============================================

function generateMultiLayerApprovalLink(
  name,
  email,
  description,
  documentType,
  attachment,
  layer,
  action
) {
  var timestamp = new Date().getTime().toString(36);
  var nameCode = name ? name.substring(0, 3).toLowerCase() : "usr";
  var projectCode = description
    ? description
        .replace(/[^a-zA-Z0-9]/g, "")
        .substring(0, 3)
        .toLowerCase()
    : "prj";
  var uniqueCode = nameCode + projectCode + timestamp;

  var params = {
    action: action || "approve",
    name: encodeURIComponent(name || ""),
    email: encodeURIComponent(email || ""),
    project: encodeURIComponent(description || ""),
    docType: encodeURIComponent(documentType || ""),
    attachment: encodeURIComponent(attachment || ""),
    layer: layer,
    code: uniqueCode,
    timestamp: new Date().getTime(),
  };

  var queryString = Object.keys(params)
    .map((key) => key + "=" + params[key])
    .join("&");

  return WEB_APP_URL + "?" + queryString;
}

// ============================================
// AUTO-TRIGGER: SEND NEXT LAYER AFTER APPROVAL
// ============================================

function sendNextApprovalAfterLevelOne() {
  Logger.log("Checking for Level Two approvals after Level One...");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:O" + sheet.getLastRow()).getValues();
  var processedCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var levelOneStatus = row[7]; // Column H
    var levelTwoStatus = row[8]; // Column I
    var overallStatus = row[14]; // Column O

    if (
      levelOneStatus === "APPROVED" &&
      (!levelTwoStatus ||
        levelTwoStatus === "" ||
        levelTwoStatus === "PENDING" ||
        levelTwoStatus === "RESUBMIT") &&
      overallStatus === "PROCESSING"
    ) {
      Logger.log("✅ Found eligible for Level Two approval: " + row[0]);

      var levelTwoEmail = row[11]; // Column L
      var name = row[0];
      var email = row[1];
      var description = row[2];
      var documentType = row[3];
      var attachment = row[4];

      if (levelTwoEmail) {
        var validationResult = validateGoogleDriveAttachmentWithType(
          attachment,
          documentType
        );
        var approvalLink = generateMultiLayerApprovalLink(
          name,
          email,
          description,
          documentType,
          attachment,
          "LEVEL_TWO",
          "approve"
        );
        var isResubmit = levelTwoStatus === "RESUBMIT";
        var emailSent = sendMultiLayerEmail(
          levelTwoEmail,
          description,
          documentType,
          attachment,
          approvalLink,
          "LEVEL_TWO",
          validationResult,
          isResubmit
        );

        if (emailSent) {
          sheet
            .getRange(i + 2, 7)
            .setNote("LEVEL_TWO_APPROVAL_SENT: " + new Date());
          Logger.log("✅ Level Two approval email sent to: " + levelTwoEmail);
          processedCount++;
          Utilities.sleep(1000);
        }
      }
    }
  }

  if (processedCount > 0) {
    Logger.log(
      "✅ Successfully sent " + processedCount + " Level Two approval emails!"
    );
  } else {
    Logger.log("No pending Level Two approvals found");
  }
}

function sendNextApprovalAfterLevelTwo() {
  Logger.log("Checking for Level Three approvals after Level Two...");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:O" + sheet.getLastRow()).getValues();
  var processedCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var levelTwoStatus = row[8]; // Column I
    var levelThreeStatus = row[9]; // Column J
    var overallStatus = row[14]; // Column O

    if (
      levelTwoStatus === "APPROVED" &&
      (!levelThreeStatus ||
        levelThreeStatus === "" ||
        levelThreeStatus === "PENDING" ||
        levelThreeStatus === "RESUBMIT") &&
      overallStatus === "PROCESSING"
    ) {
      Logger.log("✅ Found eligible for Level Three approval: " + row[0]);

      var levelThreeEmail = row[12]; // Column M
      var name = row[0];
      var email = row[1];
      var description = row[2];
      var documentType = row[3];
      var attachment = row[4];

      if (levelThreeEmail) {
        var validationResult = validateGoogleDriveAttachmentWithType(
          attachment,
          documentType
        );
        var approvalLink = generateMultiLayerApprovalLink(
          name,
          email,
          description,
          documentType,
          attachment,
          "LEVEL_THREE",
          "approve"
        );
        var isResubmit = levelThreeStatus === "RESUBMIT";
        var emailSent = sendMultiLayerEmail(
          levelThreeEmail,
          description,
          documentType,
          attachment,
          approvalLink,
          "LEVEL_THREE",
          validationResult,
          isResubmit
        );

        if (emailSent) {
          sheet
            .getRange(i + 2, 7)
            .setNote("LEVEL_THREE_APPROVAL_SENT: " + new Date());
          Logger.log(
            "✅ Level Three approval email sent to: " + levelThreeEmail
          );
          processedCount++;
          Utilities.sleep(1000);
        }
      }
    }
  }

  if (processedCount > 0) {
    Logger.log(
      "✅ Successfully sent " + processedCount + " Level Three approval emails!"
    );
  } else {
    Logger.log("No pending Level Three approvals found");
  }
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function getLayerDisplayName(layerCode) {
  var layerNames = {
    LEVEL_ONE: "Level One",
    LEVEL_TWO: "Level Two",
    LEVEL_THREE: "Level Three",
    REQUESTER_EDIT: "Requester Edit",
    LEVEL_ONE_EDIT: "Level One Edit",
    LEVEL_TWO_EDIT: "Level Two Edit",
  };
  return layerNames[layerCode] || layerCode;
}

function getGMT7Time() {
  var now = new Date();
  var timeZone = "Asia/Jakarta";
  var formattedDate = Utilities.formatDate(
    now,
    timeZone,
    "dd/MM/yyyy HH:mm:ss"
  );
  return formattedDate;
}

function showMultiLayerSummary(results, processedCount) {
  if (processedCount > 0) {
    var message =
      "Multi-Layer Approval Summary\n\nTotal processed: " +
      processedCount +
      "\n\n";

    var layerCount = {};
    results.forEach(function (result) {
      var key = result.layer + (result.isResubmit ? " (Resubmit)" : "");
      if (!layerCount[key]) layerCount[key] = 0;
      layerCount[key]++;
    });

    for (var layer in layerCount) {
      message += layer + ": " + layerCount[layer] + " emails\n";
    }

    SpreadsheetApp.getUi().alert(
      "Multi-Layer Approval Sent",
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "No Action Needed",
      "No multi-layer approvals to send.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================
// PART 2: VALIDATION & EMAIL FUNCTIONS
// Multi-Layer Reusable Approval System
// ============================================

// ============================================
// ATTACHMENT VALIDATION
// ============================================

function validateGoogleDriveAttachmentWithType(attachmentUrl, documentType) {
  try {
    if (!attachmentUrl || attachmentUrl === "") {
      return {
        valid: false,
        message: "No attachment provided",
        type: "EMPTY",
        isSharedDrive: false,
        isFolder: false,
      };
    }

    if (!attachmentUrl.includes("drive.google.com")) {
      return {
        valid: false,
        message: "Not a Google Drive link",
        type: "INVALID_URL",
        isSharedDrive: false,
        isFolder: false,
      };
    }

    var folderId = extractFolderIdFromUrl(attachmentUrl);
    if (!folderId) {
      return {
        valid: false,
        message: "Invalid Google Drive URL format",
        type: "INVALID_FORMAT",
        isSharedDrive: false,
        isFolder: false,
      };
    }

    try {
      // Try to get as folder first
      var folder = DriveApp.getFolderById(folderId);

      // It's a folder! Now validate folder contents
      Logger.log("✅ Folder detected: " + folder.getName());

      // Check if folder is in Shared Drive
      var isSharedDrive = false;
      try {
        var driveFolder = Drive.Files.get(folderId, {
          supportsAllDrives: true,
          fields: "driveId,owners",
        });
        isSharedDrive = driveFolder.driveId != null;
      } catch (e) {
        Logger.log(
          "Note: Could not check Shared Drive status: " + e.toString()
        );
        isSharedDrive = false;
      }

      // Get folder owner
      var ownerEmail = "Unknown";
      try {
        var owner = folder.getOwner();
        if (owner && owner.getEmail) {
          ownerEmail = owner.getEmail();
        }
      } catch (ownerError) {
        ownerEmail = isSharedDrive ? "Shared Drive" : "Unknown";
      }

      // List all files in folder
      var files = folder.getFiles();
      var fileList = [];
      var totalSize = 0;
      var fileCount = 0;

      while (files.hasNext() && fileCount < 50) {
        // Limit 50 files biar ga overload
        var file = files.next();
        var fileSize = file.getSize();
        totalSize += fileSize;

        fileList.push({
          name: file.getName(),
          type: file.getMimeType(),
          size: fileSize,
          sizeFormatted: formatFileSize(fileSize),
          url: file.getUrl(),
          lastUpdated: file.getLastUpdated(),
        });

        fileCount++;
      }

      // Check if folder has more files
      var hasMoreFiles = files.hasNext();

      // Document type info
      var folderInfo = "";
      if (documentType) {
        folderInfo = "Document type: " + documentType;
      }

      return {
        valid: true,
        message: "Valid Google Drive folder with " + fileCount + " file(s)",
        type: "application/vnd.google-apps.folder",
        isFolder: true,
        name: folder.getName(),
        url: folder.getUrl(),
        owner: ownerEmail,
        lastUpdated: folder.getLastUpdated(),
        folderInfo: folderInfo,
        isSharedDrive: isSharedDrive,
        driveType: isSharedDrive ? "Shared Drive" : "My Drive",
        fileCount: fileCount,
        fileList: fileList,
        totalSize: totalSize,
        totalSizeFormatted: formatFileSize(totalSize),
        hasMoreFiles: hasMoreFiles,
      };
    } catch (e) {
      Logger.log("Error accessing folder: " + e.toString());
      return {
        valid: false,
        message: "Folder not found or no access permission",
        type: "NO_ACCESS",
        isSharedDrive: false,
        isFolder: false,
      };
    }
  } catch (error) {
    Logger.log("Validation error: " + error.toString());
    return {
      valid: false,
      message: "Validation error: " + error.message,
      type: "VALIDATION_ERROR",
      isSharedDrive: false,
      isFolder: false,
    };
  }
}

//function extractFileIdFromUrl(url) {
//  var patterns = [
//    /\/d\/([a-zA-Z0-9_-]+)/,
//    /id=([a-zA-Z0-9_-]+)/,
//    /\/file\/d\/([a-zA-Z0-9_-]+)/,
//    /\/open\?id=([a-zA-Z0-9_-]+)/
//  ];
//
//  for (var i = 0; i < patterns.length; i++) {
//    var match = url.match(patterns[i]);
//    if (match && match[1]) {
//      return match[1];
//    }
//  }
//  return null;
//}

function extractFolderIdFromUrl(url) {
  var patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/, // Standard folder URL
    /\/drive\/folders\/([a-zA-Z0-9_-]+)/, // Drive folder URL
    /id=([a-zA-Z0-9_-]+)/, // Query param
    /\/d\/([a-zA-Z0-9_-]+)/, // File ID format (fallback)
    /\/file\/d\/([a-zA-Z0-9_-]+)/, // File format (fallback)
    /\/open\?id=([a-zA-Z0-9_-]+)/, // Open format (fallback)
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match && match[1]) {
      return match[1];
    }
  }
  return null;
}

function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  var k = 1024;
  var sizes = ["Bytes", "KB", "MB", "GB"];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

// ============================================
// EMAIL SENDING FUNCTIONS
// ============================================

function sendMultiLayerEmail(
  recipientEmail,
  description,
  documentType,
  attachment,
  approvalLink,
  layer,
  validationResult,
  isResubmit
) {
  try {
    if (!recipientEmail || recipientEmail.indexOf("@") === -1) {
      Logger.log("Invalid email: " + recipientEmail);
      return false;
    }

    var layerDisplayNames = {
      LEVEL_ONE: "Level One",
      LEVEL_TWO: "Level Two",
      LEVEL_THREE: "Level Three",
      LEVEL_ONE_EDIT: "Level One (Edit & Resubmit)",
      LEVEL_TWO_EDIT: "Level Two (Edit & Resubmit)",
    };

    var layerDisplay = layerDisplayNames[layer] || layer;
    var companyName = "Atreus Global";

    // Subject line berbeda untuk resubmit
    var subject = isResubmit
      ? "🔄 Resubmit Required - " + description + " [" + layerDisplay + "]"
      : "Approval Request - " + description + " [" + layerDisplay + "]";

    // Document type badge
    var docTypeBadge = "";
    if (documentType && documentType !== "") {
      var docTypeColors = {
        ICC: "#3B82F6",
        Quotation: "#10B981",
        Proposal: "#F59E0B",
      };
      var badgeColor = docTypeColors[documentType] || "#6B7280";
      docTypeBadge =
        '<div style="display: inline-block; background: ' +
        badgeColor +
        '; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-bottom: 10px;">' +
        documentType +
        "</div>";
    }

    // Attachment section
    var attachmentSection = "";
    if (attachment && attachment !== "") {
      if (validationResult.valid) {
        var driveTypeInfo = validationResult.isSharedDrive
          ? '<p><strong>Location:</strong> <span style="color: #3B82F6;">📁 Shared Drive</span></p>'
          : "<p><strong>Location:</strong> My Drive</p>";

        // Check if it's a folder
        if (validationResult.isFolder) {
          // FOLDER VIEW - Show file list
          var fileListHtml = "";
          if (
            validationResult.fileList &&
            validationResult.fileList.length > 0
          ) {
            fileListHtml =
              '<div style="margin-top: 15px;"><h4 style="color: #2E8B57; margin-bottom: 10px;">📂 Folder Contents:</h4><ul style="list-style: none; padding: 0; margin: 0;">';

            validationResult.fileList.forEach(function (file) {
              var fileIcon = "📄";
              if (file.type.includes("image")) fileIcon = "🖼️";
              else if (file.type.includes("pdf")) fileIcon = "📕";
              else if (file.type.includes("document")) fileIcon = "📝";
              else if (file.type.includes("spreadsheet")) fileIcon = "📊";

              fileListHtml +=
                '<li style="background: #f8f9fa; padding: 10px; margin-bottom: 5px; border-radius: 5px; border-left: 3px solid #90EE90;"><strong>' +
                fileIcon +
                " " +
                file.name +
                '</strong><br><span style="font-size: 11px; color: #666;">Type: ' +
                file.type.split("/").pop() +
                " | Size: " +
                file.sizeFormatted +
                "</span></li>";
            });

            if (validationResult.hasMoreFiles) {
              fileListHtml +=
                '<li style="color: #666; font-style: italic; padding: 5px;">...and more files (showing first 50)</li>';
            }

            fileListHtml += "</ul></div>";
          }

          attachmentSection =
            '<div class="attachment-box" style="background: linear-gradient(135deg, #f0f8ff 0%, #e6f2ff 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #90EE90;"><h3 style="color: #2E8B57;">📁 Folder Attachment (Validated)</h3>' +
            docTypeBadge +
            "<p><strong>Folder Name:</strong> " +
            validationResult.name +
            "</p><p><strong>Total Files:</strong> " +
            validationResult.fileCount +
            " file(s)</p><p><strong>Total Size:</strong> " +
            validationResult.totalSizeFormatted +
            "</p>" +
            driveTypeInfo +
            '<p><strong>Status:</strong> <span style="color: #2E8B57;">' +
            validationResult.message +
            "</span></p>" +
            fileListHtml +
            '<p style="margin-top: 15px;"><a href="' +
            attachment +
            '" target="_blank" style="color: white; background: #326BC6; padding: 10px 20px; border-radius: 5px; text-decoration: none; font-weight: 600; display: inline-block;">📂 Open Folder</a></p></div>';
        } else {
          // SINGLE FILE VIEW (original)
          attachmentSection =
            '<div class="attachment-box" style="background: linear-gradient(135deg, #f0f8ff 0%, #e6f2ff 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #90EE90;"><h3 style="color: #2E8B57;">Attachment (Validated)</h3>' +
            docTypeBadge +
            "<p><strong>File:</strong> " +
            validationResult.name +
            "</p><p><strong>Type:</strong> " +
            validationResult.type +
            "</p><p><strong>Size:</strong> " +
            validationResult.sizeFormatted +
            "</p>" +
            driveTypeInfo +
            '<p><strong>Status:</strong> <span style="color: #2E8B57;">' +
            validationResult.message +
            '</span></p><p><a href="' +
            attachment +
            '" target="_blank" style="color: #326BC6; font-weight: 600;">📎 View Document</a></p></div>';
        }
      } else {
        attachmentSection =
          '<div class="attachment-box" style="background: linear-gradient(135deg, #fff0f0 0%, #ffe6e6 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #FF6B6B;"><h3 style="color: #DC143C;">Attachment (Validation Failed)</h3>' +
          docTypeBadge +
          '<p><strong>Status:</strong> <span style="color: #DC143C;">' +
          validationResult.message +
          '</span></p><p><a href="' +
          attachment +
          '" target="_blank" style="color: #326BC6;">Check Link</a></p><p style="font-size: 12px; color: #666;"><em>Please ensure the Google Drive link is correct and accessible</em></p></div>';
      }
    }

    var progressBar = getProgressBar(layer);

    // Generate rejection link
    var rejectLink = generateMultiLayerApprovalLink(
      "",
      "",
      description,
      documentType,
      attachment,
      layer,
      "reject"
    );

    // Resubmit notice (jika ini resubmit)
    var resubmitNotice = "";
    if (isResubmit) {
      resubmitNotice =
        '<div style="background: #FFF4E6; border-left: 4px solid #F59E0B; padding: 15px; border-radius: 8px; margin: 20px 0;"><strong style="color: #F59E0B;">🔄 RESUBMIT REQUIRED</strong><p style="margin: 8px 0 0 0; color: #92400E;">This document was previously rejected and has been revised. Please review the changes and approve or reject again.</p></div>';
    }

    var htmlBody =
      '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;line-height:1.6;color:#333;background:#f6f9fc;margin:0;padding:0}.container{max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1)}.header{background:linear-gradient(135deg,#326BC6 0%,#183460 100%);padding:30px;text-align:center;color:white}.progress-track{background:#f8f9fa;padding:30px 20px;margin:0}.progress-steps{display:flex;justify-content:center;align-items:center;gap:40px;position:relative}.progress-step{text-align:center;position:relative}.step-number{width:40px;height:40px;border-radius:50%;margin:0 auto 8px;font-weight:600;display:flex;align-items:center;justify-content:center;font-size:16px}.step-active{background:#326BC6;color:white}.step-completed{background:#90EE90;color:#333}.step-pending{background:#e0e0e0;color:#666}.step-label{font-size:12px;font-weight:500}.content{padding:30px}.button{display:inline-block;padding:15px 30px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white !important;text-decoration:none;border-radius:8px;margin:20px 10px;font-weight:600;font-size:16px;border:none;cursor:pointer;transition:all 0.3s ease}.button:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(50,107,198,0.3)}.button-reject{background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white !important;padding:15px 30px;text-decoration:none;border-radius:8px;margin:20px 10px;font-weight:600;font-size:16px;border:none;cursor:pointer;transition:all 0.3s ease}.button-reject:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(220,38,38,0.3)}.button-container{text-align:center;margin:20px 0}.info-box{background:#f8f9fa;padding:15px;border-radius:5px;margin:15px 0;border-left:4px solid #326BC6}.footer{margin-top:30px;padding:20px;background:#f8f9fa;text-align:center;font-size:12px;color:#666}</style></head><body><div class="container"><div class="header"><h1>Multi-Layer Approval ' +
      (isResubmit ? "Resubmit" : "Required") +
      "</h1><p>Current Stage: " +
      layerDisplay +
      '</p></div><div class="progress-track">' +
      progressBar +
      '</div><div class="content">' +
      resubmitNotice +
      "<p>Hello,</p><p>This project requires your approval at the <strong>" +
      layerDisplay +
      '</strong> level:</p><div class="info-box"><h3>' +
      description +
      "</h3><p><strong>Current Stage:</strong> " +
      layerDisplay +
      "</p><p><strong>Date:</strong> " +
      getGMT7Time() +
      "</p></div>" +
      attachmentSection +
      '<p>Please choose your decision:</p><div class="button-container"><a href="' +
      approvalLink +
      '" class="button" style="color: white !important;">✓ APPROVE</a><a href="' +
      rejectLink +
      '" class="button-reject" style="color: white !important;">✗ REJECT & SEND BACK</a></div><p style="text-align: center; font-size: 12px; color: #666;"><em>Link expires in 7 days</em></p></div><div class="footer"><p>This is an automated email. Please do not reply.</p><p>© ' +
      new Date().getFullYear() +
      " " +
      companyName +
      ". All rights reserved.</p></div></div></body></html>";

    var plainBody =
      "MULTI-LAYER APPROVAL REQUEST" +
      (isResubmit ? " (RESUBMIT)" : "") +
      "\n\n" +
      "Project: " +
      description +
      "\n" +
      "Document Type: " +
      documentType +
      "\n" +
      "Current Stage: " +
      layerDisplay +
      "\n\n";

    if (attachment && attachment !== "") {
      plainBody +=
        "Attachment: " +
        attachment +
        "\n" +
        "Attachment Status: " +
        validationResult.message +
        "\n\n";
    }

    plainBody +=
      "Approve: " +
      approvalLink +
      "\n" +
      "Reject: " +
      rejectLink +
      "\n\n" +
      "Thank you.\n\n" +
      "This is an automated email from " +
      companyName +
      ".";

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody,
    });

    Logger.log(
      "✅ Email sent to: " +
        recipientEmail +
        " for layer: " +
        layerDisplay +
        (isResubmit ? " (Resubmit)" : "")
    );
    return true;
  } catch (error) {
    Logger.log("❌ Error sending email: " + error.toString());
    return false;
  }
}

function getProgressBar(currentLayer) {
  var steps = [
    {
      number: 1,
      label: "Level One",
      status:
        currentLayer === "LEVEL_ONE" || currentLayer === "LEVEL_ONE_EDIT"
          ? "active"
          : [
              "LEVEL_TWO",
              "LEVEL_TWO_EDIT",
              "LEVEL_THREE",
              "COMPLETED",
            ].includes(currentLayer)
          ? "completed"
          : "pending",
    },
    {
      number: 2,
      label: "Level Two",
      status:
        currentLayer === "LEVEL_TWO" || currentLayer === "LEVEL_TWO_EDIT"
          ? "active"
          : ["LEVEL_THREE", "COMPLETED"].includes(currentLayer)
          ? "completed"
          : "pending",
    },
    {
      number: 3,
      label: "Level Three",
      status:
        currentLayer === "LEVEL_THREE"
          ? "active"
          : currentLayer === "COMPLETED"
          ? "completed"
          : "pending",
    },
  ];

  var progressHtml = '<div class="progress-steps">';

  steps.forEach(function (step) {
    progressHtml +=
      '<div class="progress-step"><div class="step-number step-' +
      step.status +
      '">' +
      step.number +
      '</div><div class="step-label">' +
      step.label +
      "</div></div>";
  });

  progressHtml += "</div>";
  return progressHtml;
}

// ============================================
// SEND BACK NOTIFICATION AFTER REJECTION
// ============================================

function sendSendBackNotification(
  recipientEmail,
  description,
  documentType,
  layer,
  rejectionNote,
  rejectorName
) {
  try {
    if (!recipientEmail || recipientEmail.indexOf("@") === -1) {
      Logger.log("Invalid email: " + recipientEmail);
      return false;
    }

    var layerDisplayNames = {
      LEVEL_ONE: "Level One",
      LEVEL_TWO: "Level Two",
      LEVEL_THREE: "Level Three",
    };

    var rejectedAtLayer = layerDisplayNames[layer] || layer;
    var companyName = "Atreus Global";

    var subject = "📝 Document Revision Required - " + description;

    var htmlBody =
      '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;line-height:1.6;color:#333;background:#f6f9fc;margin:0;padding:0}.container{max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1)}.header{background:linear-gradient(135deg,#F59E0B 0%,#D97706 100%);padding:30px;text-align:center;color:white}.content{padding:30px}.rejection-box{background:#FEF3C7;border-left:4px solid #F59E0B;padding:20px;border-radius:8px;margin:20px 0}.info-box{background:#f8f9fa;padding:15px;border-radius:5px;margin:15px 0;border-left:4px solid #F59E0B}.footer{margin-top:30px;padding:20px;background:#f8f9fa;text-align:center;font-size:12px;color:#666}</style></head><body><div class="container"><div class="header"><h1>📝 Revision Required</h1><p>Document has been sent back for editing</p></div><div class="content"><p>Hello,</p><p>The document <strong>' +
      description +
      "</strong> has been rejected at <strong>" +
      rejectedAtLayer +
      '</strong> and sent back to you for revision.</p><div class="info-box"><p><strong>Project:</strong> ' +
      description +
      "</p><p><strong>Document Type:</strong> " +
      documentType +
      "</p><p><strong>Rejected by:</strong> " +
      (rejectorName || rejectedAtLayer + " Approver") +
      "</p><p><strong>Date:</strong> " +
      getGMT7Time() +
      '</p></div><div class="rejection-box"><h3 style="color:#92400E;margin-top:0;">Rejection Feedback:</h3><p style="margin:0;color:#78350F;">' +
      rejectionNote.replace(/\n/g, "<br>") +
      '</p></div><p><strong>Next Steps:</strong></p><ol><li>Review the feedback above</li><li>Make necessary revisions to the document</li><li>Update the attachment link in the spreadsheet</li><li>Check the "Send" checkbox to resubmit for approval</li></ol><p style="color:#92400E;"><em>💡 The approval will automatically continue from where it was rejected after you resubmit.</em></p></div><div class="footer"><p>This is an automated email. Please do not reply.</p><p>© ' +
      new Date().getFullYear() +
      " " +
      companyName +
      ". All rights reserved.</p></div></div></body></html>";

    var plainBody =
      "DOCUMENT REVISION REQUIRED\n\n" +
      "Project: " +
      description +
      "\n" +
      "Document Type: " +
      documentType +
      "\n" +
      "Rejected at: " +
      rejectedAtLayer +
      "\n" +
      "Rejected by: " +
      (rejectorName || rejectedAtLayer + " Approver") +
      "\n" +
      "Date: " +
      getGMT7Time() +
      "\n\n" +
      "REJECTION FEEDBACK:\n" +
      rejectionNote +
      "\n\n" +
      "NEXT STEPS:\n" +
      "1. Review the feedback above\n" +
      "2. Make necessary revisions\n" +
      "3. Update the attachment link in the spreadsheet\n" +
      "4. Check the 'Send' checkbox to resubmit\n\n" +
      "This is an automated email from " +
      companyName +
      ".";

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody,
    });

    Logger.log("✅ Send back notification sent to: " + recipientEmail);
    return true;
  } catch (error) {
    Logger.log("❌ Error sending send back notification: " + error.toString());
    return false;
  }
}

function sendAdminNotification(message, subject) {
  var ADMIN_EMAIL = "mhmdfdln14@gmail.com"; // Ganti dengan email admin lu
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject || "Approval System Notification",
      body: message,
    });
    Logger.log("✅ Admin notification sent");
  } catch (error) {
    Logger.log("❌ Failed to send admin notification: " + error.toString());
  }
}

// ============================================
// PART 3: WEB APP HANDLERS (APPROVAL & REJECTION)
// Multi-Layer Reusable Approval System
// ============================================

// ============================================
// WEB APP ENTRY POINTS
// ============================================

function doGet(e) {
  try {
    Logger.log("Web App accessed");
    Logger.log("Parameters: " + JSON.stringify(e.parameter));

    var params = e.parameter || {};
    var action = params.action;

    if (action === "approve") {
      return handleMultiLayerApproval(params);
    } else if (action === "reject") {
      return handleMultiLayerRejection(params);
    } else if (action === "submit_rejection") {
      // Handle rejection submission via GET
      return handleRejectionSubmission(params);
    } else if (action === "rejection_success") {
      return showRejectionSuccessPage(params);
    }

    return createErrorPage("Invalid request - missing action parameter");
  } catch (error) {
    Logger.log("❌ Error in doGet: " + error.toString());
    return createErrorPage("System error: " + error.message);
  }
}

function doPost(e) {
  try {
    Logger.log("POST request received");

    var params = {};

    if (e.postData && e.postData.contents) {
      var contents = e.postData.contents;
      var formData = contents.split("&");

      for (var i = 0; i < formData.length; i++) {
        var pair = formData[i].split("=");
        if (pair.length === 2) {
          params[decodeURIComponent(pair[0])] = decodeURIComponent(
            pair[1].replace(/\+/g, " ")
          );
        }
      }
    }

    Logger.log("Processed POST params: " + JSON.stringify(params));

    var action = params.action;

    if (action === "submit_rejection") {
      return handleRejectionSubmission(params);
    }

    return createErrorPage("Invalid POST request");
  } catch (error) {
    Logger.log("❌ Error in doPost: " + error.toString());
    return createErrorPage("System error: " + error.message);
  }
}

// ============================================
// APPROVAL HANDLER
// ============================================

function handleMultiLayerApproval(params) {
  try {
    Logger.log("=== APPROVAL REQUEST RECEIVED ===");

    var name = params.name ? decodeURIComponent(params.name) : "";
    var email = params.email ? decodeURIComponent(params.email) : "";
    var project = params.project ? decodeURIComponent(params.project) : "";
    var documentType = params.docType ? decodeURIComponent(params.docType) : "";
    var attachment = params.attachment
      ? decodeURIComponent(params.attachment)
      : "";
    var layer = params.layer ? decodeURIComponent(params.layer) : "";
    var code = params.code || "";
    var timestamp = parseInt(params.timestamp) || 0;

    Logger.log("Approving: " + project + " at " + layer);

    // Check link expiration (7 days)
    var now = new Date().getTime();
    var sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000;
    if (timestamp < sevenDaysAgo) {
      return createErrorPage(
        "Approval link has expired (7 days). Please request a new one."
      );
    }

    if (!project || !layer) {
      return createErrorPage("Missing required approval data");
    }

    // Update approval status
    var updated = updateMultiLayerApprovalStatus(
      name,
      email,
      project,
      layer,
      code
    );

    if (updated) {
      Logger.log("✅ Approval updated successfully");

      // Auto-trigger next layer approval
      try {
        if (layer === "LEVEL_ONE") {
          Utilities.sleep(2000);
          sendNextApprovalAfterLevelOne();
        } else if (layer === "LEVEL_TWO") {
          Utilities.sleep(2000);
          sendNextApprovalAfterLevelTwo();
        }
      } catch (nextError) {
        Logger.log("⚠️ Next approval trigger failed: " + nextError.toString());
      }

      return createSuccessPage(
        name,
        email,
        project,
        documentType,
        attachment,
        layer,
        code
      );
    } else {
      return createErrorPage(
        "Approval failed - data not found or already processed"
      );
    }
  } catch (error) {
    Logger.log("❌ Error in handleMultiLayerApproval: " + error.toString());
    return createErrorPage("System error during approval: " + error.message);
  }
}

// ============================================
// UPDATE APPROVAL STATUS (CORE LOGIC)
// ============================================

function updateMultiLayerApprovalStatus(name, email, project, layer, code) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;

    var data = sheet.getRange("A2:O" + lastRow).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowName = row[0];
      var rowEmail = row[1];
      var rowProject = row[2];
      var rowStatus = row[14]; // Column O - Overall Status

      if (!rowName && !rowEmail && !rowProject) continue;

      // STRICT MATCHING
      var projectMatch =
        rowProject &&
        project &&
        rowProject.toString().trim().toLowerCase() ===
          project.toString().trim().toLowerCase();

      var isEligible = rowStatus === "PROCESSING" || rowStatus === "ACTIVE";

      if (projectMatch && isEligible) {
        Logger.log("✅ MATCH FOUND at row " + (i + 2));

        var columnIndex = getLayerColumnIndex(layer);
        if (columnIndex === -1) {
          Logger.log("❌ Invalid layer: " + layer);
          return false;
        }

        var currentStatus = sheet.getRange(i + 2, columnIndex).getValue();

        // Prevent double approval
        if (currentStatus === "APPROVED") {
          Logger.log("⚠️ Already approved - skipping");
          return false;
        }

        // SET APPROVED
        sheet.getRange(i + 2, columnIndex).setValue("APPROVED");
        sheet.getRange(i + 2, columnIndex).setBackground("#90EE90");
        sheet
          .getRange(i + 2, columnIndex)
          .setNote(
            "Approved by " +
              getLayerDisplayName(layer) +
              " - " +
              getGMT7Time() +
              " - Code: " +
              code
          );

        // Clear "Current Editor" field
        sheet.getRange(i + 2, 14).setValue(""); // Column N - Current Editor

        // Check if all layers approved
        var levelOneStatus = sheet.getRange(i + 2, 8).getValue(); // Column H
        var levelTwoStatus = sheet.getRange(i + 2, 9).getValue(); // Column I
        var levelThreeStatus = sheet.getRange(i + 2, 10).getValue(); // Column J

        if (
          levelOneStatus === "APPROVED" &&
          levelTwoStatus === "APPROVED" &&
          levelThreeStatus === "APPROVED"
        ) {
          sheet.getRange(i + 2, 15).setValue("COMPLETED"); // Column O
          sheet.getRange(i + 2, 15).setBackground("#90EE90");
          Logger.log("🎉 ALL LAYERS APPROVED - COMPLETED");
        } else {
          sheet.getRange(i + 2, 15).setValue("PROCESSING");
          sheet.getRange(i + 2, 15).setBackground("#FFF2CC");
        }

        Logger.log("✅ Approval recorded for: " + layer);
        return true;
      }
    }

    Logger.log("❌ No matching data found");
    return false;
  } catch (error) {
    Logger.log("❌ Error updating approval status: " + error.toString());
    return false;
  }
}

function getLayerColumnIndex(layer) {
  var layerColumns = {
    LEVEL_ONE: 8, // Column H
    LEVEL_TWO: 9, // Column I
    LEVEL_THREE: 10, // Column J
    LEVEL_ONE_EDIT: 8, // Same as LEVEL_ONE (for resubmit)
    LEVEL_TWO_EDIT: 9, // Same as LEVEL_TWO (for resubmit)
  };
  return layerColumns[layer] || -1;
}

// ============================================
// REJECTION HANDLER
// ============================================

function handleMultiLayerRejection(params) {
  try {
    Logger.log("=== REJECTION REQUEST RECEIVED ===");

    var name = params.name ? decodeURIComponent(params.name) : "";
    var email = params.email ? decodeURIComponent(params.email) : "";
    var project = params.project ? decodeURIComponent(params.project) : "";
    var documentType = params.docType ? decodeURIComponent(params.docType) : "";
    var attachment = params.attachment
      ? decodeURIComponent(params.attachment)
      : "";
    var layer = params.layer ? decodeURIComponent(params.layer) : "";
    var code = params.code || "";
    var timestamp = parseInt(params.timestamp) || 0;

    Logger.log("Rejection for: " + project + " at " + layer);

    // Check expiration
    var now = new Date().getTime();
    var sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000;
    if (timestamp < sevenDaysAgo) {
      return createErrorPage(
        "Rejection link has expired (7 days). Please request a new one."
      );
    }

    if (!project || !layer) {
      return createErrorPage("Missing required rejection data");
    }

    return createRejectionForm(
      name,
      email,
      project,
      documentType,
      attachment,
      layer,
      code
    );
  } catch (error) {
    Logger.log("❌ Error in handleMultiLayerRejection: " + error.toString());
    return createErrorPage("System error during rejection: " + error.message);
  }
}

function handleRejectionSubmission(params) {
  try {
    Logger.log("=== REJECTION SUBMISSION RECEIVED ===");

    var name = params.name || "";
    var email = params.email || "";
    var project = params.project || "";
    var documentType = params.docType || "";
    var attachment = params.attachment || "";
    var layer = params.layer || "";
    var code = params.code || "";
    var rejectionNote = params.rejectionNote || "";

    Logger.log("Processing rejection: " + project + " at " + layer);

    if (!rejectionNote || rejectionNote.trim() === "") {
      return createErrorPage("Rejection note is required.");
    }

    // Update spreadsheet with rejection
    var updated = updateMultiLayerRejectionStatus(
      name,
      email,
      project,
      layer,
      code,
      rejectionNote
    );

    if (updated) {
      Logger.log("✅ Rejection recorded successfully");

      // Send notification ke yang harus edit
      try {
        sendRejectionAndSendBackNotifications(
          name,
          email,
          project,
          documentType,
          attachment,
          layer,
          rejectionNote
        );
      } catch (notifyError) {
        Logger.log("⚠️ Notification failed: " + notifyError.toString());
      }

      return createRejectionSuccessPage(
        name,
        email,
        project,
        documentType,
        attachment,
        layer,
        rejectionNote
      );
    } else {
      Logger.log("❌ Rejection failed");
      return createErrorPage(
        "Rejection failed - data not found in spreadsheet"
      );
    }
  } catch (error) {
    Logger.log("❌ Error in handleRejectionSubmission: " + error.toString());
    return createErrorPage("System error during rejection: " + error.message);
  }
}

// ============================================
// UPDATE REJECTION STATUS (CORE LOGIC - REUSABLE)
// ============================================

function updateMultiLayerRejectionStatus(
  name,
  email,
  project,
  layer,
  code,
  rejectionNote
) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;

    var data = sheet.getRange("A2:O" + lastRow).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowProject = row[2];
      var rowStatus = row[14]; // Column O

      if (!rowProject) continue;

      var projectMatch =
        rowProject.toString().trim().toLowerCase() ===
        project.toString().trim().toLowerCase();
      var isEligible = rowStatus === "PROCESSING" || rowStatus === "ACTIVE";

      if (projectMatch && isEligible) {
        Logger.log("✅ MATCH FOUND at row " + (i + 2) + " for rejection");

        var columnIndex = getLayerColumnIndex(layer);
        if (columnIndex === -1) return false;

        // Check current status
        var currentStatus = sheet.getRange(i + 2, columnIndex).getValue();
        if (currentStatus === "REJECTED") {
          Logger.log("⚠️ Already rejected - skipping");
          return false;
        }

        // SET REJECTED and clear next levels
        sheet.getRange(i + 2, columnIndex).setValue("REJECTED");
        sheet.getRange(i + 2, columnIndex).setBackground("#FF6B6B");
        sheet
          .getRange(i + 2, columnIndex)
          .setNote(
            "Rejected by " +
              getLayerDisplayName(layer) +
              " - " +
              getGMT7Time() +
              "\nReason: " +
              rejectionNote
          );

        // Reset next levels to PENDING
        if (layer === "LEVEL_ONE") {
          // Clear Level Two and Three
          sheet.getRange(i + 2, 9).setValue("PENDING"); // Level Two
          sheet.getRange(i + 2, 9).setBackground(null);
          sheet.getRange(i + 2, 10).setValue("PENDING"); // Level Three
          sheet.getRange(i + 2, 10).setBackground(null);
        } else if (layer === "LEVEL_TWO") {
          // Clear Level Three
          sheet.getRange(i + 2, 10).setValue("PENDING"); // Level Three
          sheet.getRange(i + 2, 10).setBackground(null);
        }

        // SEND BACK LOGIC with enhanced status handling
        if (layer === "LEVEL_ONE") {
          // Rejected at Level One -> send back to REQUESTER
          sheet.getRange(i + 2, 14).setValue("REQUESTER"); // Column N - Current Editor
          sheet.getRange(i + 2, 15).setValue("EDITING"); // Column O - Overall Status
          sheet.getRange(i + 2, 15).setBackground("#FFE0B2");
          sheet.getRange(i + 2, 8).setValue("REJECTED"); // Ensure Level One shows REJECTED
          sheet.getRange(i + 2, 8).setBackground("#FF6B6B");
        } else if (layer === "LEVEL_TWO") {
          // Rejected at Level Two -> send back to LEVEL ONE
          sheet.getRange(i + 2, 14).setValue("LEVEL_ONE"); // Column N - Current Editor
          sheet.getRange(i + 2, 15).setValue("EDITING"); // Column O - Overall Status
          sheet.getRange(i + 2, 15).setBackground("#FFE0B2");
          sheet.getRange(i + 2, 9).setValue("REJECTED"); // Ensure Level Two shows REJECTED
          sheet.getRange(i + 2, 9).setBackground("#FF6B6B");
          sheet.getRange(i + 2, 8).setValue("EDITING"); // Set Level One to EDITING
          sheet.getRange(i + 2, 8).setBackground("#FFE0B2");
        } else if (layer === "LEVEL_THREE") {
          // Rejected at Level Three -> send back to LEVEL TWO
          sheet.getRange(i + 2, 14).setValue("LEVEL_TWO"); // Column N - Current Editor
          sheet.getRange(i + 2, 15).setValue("EDITING"); // Column O - Overall Status
          sheet.getRange(i + 2, 15).setBackground("#FFE0B2");
          sheet.getRange(i + 2, 10).setValue("REJECTED"); // Ensure Level Three shows REJECTED
          sheet.getRange(i + 2, 10).setBackground("#FF6B6B");
          sheet.getRange(i + 2, 9).setValue("EDITING"); // Set Level Two to EDITING
          sheet.getRange(i + 2, 9).setBackground("#FFE0B2");
        }

        Logger.log("✅ Rejection recorded - sent back to: " + sendBackTo);

        // Store info untuk notification
        row.sendBackTo = sendBackTo;
        row.rowIndex = i + 2;

        return row; // Return row data untuk notification
      }
    }

    Logger.log("❌ No matching data found for rejection");
    return false;
  } catch (error) {
    Logger.log("❌ Error updating rejection status: " + error.toString());
    return false;
  }
}

// ============================================
// SEND NOTIFICATIONS AFTER REJECTION
// ============================================

function sendRejectionAndSendBackNotifications(
  name,
  email,
  project,
  documentType,
  attachment,
  layer,
  rejectionNote
) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:O" + sheet.getLastRow()).getValues();

    // Find the row
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowProject = row[2];

      if (
        rowProject &&
        rowProject.toString().trim().toLowerCase() ===
          project.toString().trim().toLowerCase()
      ) {
        var requesterEmail = row[1]; // Column B
        var levelOneEmail = row[10]; // Column K
        var levelTwoEmail = row[11]; // Column L
        var currentEditor = row[13]; // Column N

        var sendToEmail = "";
        var sendToName = "";

        // Determine siapa yang harus di-notify untuk edit
        if (layer === "LEVEL_ONE") {
          // Rejected at Level One -> notify REQUESTER
          sendToEmail = requesterEmail;
          sendToName = row[0]; // Name from Column A
        } else if (layer === "LEVEL_TWO") {
          // Rejected at Level Two -> notify LEVEL ONE
          sendToEmail = levelOneEmail;
          sendToName = "Level One Approver";
        } else if (layer === "LEVEL_THREE") {
          // Rejected at Level Three -> notify LEVEL TWO
          sendToEmail = levelTwoEmail;
          sendToName = "Level Two Approver";
        }

        if (sendToEmail) {
          sendSendBackNotification(
            sendToEmail,
            project,
            documentType,
            layer,
            rejectionNote,
            sendToName
          );
          Logger.log("✅ Send back notification sent to: " + sendToEmail);
        }

        break;
      }
    }
  } catch (error) {
    Logger.log("❌ Error sending notifications: " + error.toString());
  }
}

// ============================================
// SUCCESS & ERROR PAGES
// ============================================

function createSuccessPage(
  name,
  email,
  project,
  documentType,
  attachment,
  layer,
  code
) {
  var layerDisplay = getLayerDisplayName(layer);

  var html =
    '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%;margin:0 auto}.success-icon{font-size:64px;margin-bottom:20px;color:#10B981}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#10B981}.details{background:#f8fafc;padding:20px;border-radius:10px;margin:20px 0;text-align:left}.detail-item{margin-bottom:10px;display:flex}.detail-label{font-weight:600;color:#326BC6;min-width:120px}.detail-value{color:#334155;flex:1}</style></head><body><div class="container"><div class="success-icon">✓</div><h1 class="title">Approval Successful!</h1><p style="color:#10B981;margin-bottom:25px;">' +
    layerDisplay +
    ' has been approved</p><div class="details"><div class="detail-item"><span class="detail-label">Project:</span><span class="detail-value">' +
    (project || "N/A") +
    '</span></div><div class="detail-item"><span class="detail-label">Layer:</span><span class="detail-value">' +
    layerDisplay +
    '</span></div><div class="detail-item"><span class="detail-label">Date:</span><span class="detail-value">' +
    getGMT7Time() +
    '</span></div></div><p style="font-size:13px;color:#64748b;margin-top:20px;">You can safely close this page.</p></div></body></html>';

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function createRejectionForm(
  name,
  email,
  project,
  documentType,
  attachment,
  layer,
  code
) {
  var layerDisplay = getLayerDisplayName(layer);

  // Create HTML template with proper Google Apps Script integration
  var template = HtmlService.createTemplateFromFile("RejectionForm");
  template.name = name || "";
  template.email = email || "";
  template.project = project || "";
  template.documentType = documentType || "";
  template.attachment = attachment || "";
  template.layer = layer || "";
  template.layerDisplay = layerDisplay;
  template.code = code || "";

  return template
    .evaluate()
    .setTitle("Rejection Form")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Create separate HTML file content (inline version since we can't create files dynamically)
function createRejectionForm(
  name,
  email,
  project,
  documentType,
  attachment,
  layer,
  code
) {
  var layerDisplay = getLayerDisplayName(layer);
  var isEditor = layer === "LEVEL_ONE" || layer === "LEVEL_TWO";

  var docTypeBadge = "";
  if (documentType && documentType !== "") {
    var docTypeColors = {
      ICC: "#3B82F6",
      Quotation: "#10B981",
      Proposal: "#F59E0B",
    };
    var badgeColor = docTypeColors[documentType] || "#6B7280";
    docTypeBadge =
      '<div style="display: inline-block; background: ' +
      badgeColor +
      '; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; margin: 10px 0;">' +
      documentType +
      "</div>";
  }

  // Generate submit rejection link (sama kayak approval link)
  var submitRejectionLink = generateRejectionSubmitLink(
    name,
    email,
    project,
    documentType,
    attachment,
    layer,
    code,
    "PLACEHOLDER_REJECTION_NOTE"
  );

  // PAKE METODE YANG SAMA KAYAK APPROVAL - SIMPLE HTML + DIRECT LINK
  var html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    @import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");
    body {
      font-family: "Inter", sans-serif;
      text-align: center;
      padding: 20px;
      background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%);
      color: white;
      margin: 0;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .container {
      background: white;
      color: #333;
      padding: 40px 30px;
      border-radius: 16px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.15);
      max-width: 600px;
      width: 90%;
      margin: 0 auto;
    }
    .error-icon {
      font-size: 64px;
      margin-bottom: 20px;
      color: #dc2626;
      font-weight: 700;
    }
    .title {
      font-weight: 700;
      font-size: 28px;
      margin-bottom: 8px;
      color: #dc2626;
    }
    .subtitle {
      font-weight: 500;
      font-size: 16px;
      color: #dc2626;
      margin-bottom: 25px;
    }
    .rejection-form {
      background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
      padding: 20px;
      border-radius: 10px;
      margin: 20px 0;
      border-left: 4px solid #dc2626;
      text-align: left;
    }
    .form-group {
      margin-bottom: 15px;
    }
    .form-label {
      font-weight: 600;
      color: #dc2626;
      margin-bottom: 5px;
      display: block;
    }
    .form-textarea {
      width: 100%;
      padding: 10px;
      border: 1px solid #e5e7eb;
      border-radius: 5px;
      font-family: inherit;
      font-size: 14px;
      min-height: 100px;
      resize: vertical;
      box-sizing: border-box;
    }
    .button {
      display: inline-block;
      padding: 15px 30px;
      background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%);
      color: white !important;
      text-decoration: none;
      border-radius: 8px;
      margin: 10px;
      font-weight: 600;
      font-size: 16px;
      border: none;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    .button:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 12px rgba(220,38,38,0.3);
    }
    .button:disabled {
      opacity: 0.6;
      cursor: not-allowed;
      transform: none;
    }
    .details {
      background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
      padding: 20px;
      border-radius: 10px;
      margin: 20px 0;
      text-align: left;
      border-left: 4px solid #dc2626;
    }
    .detail-item {
      margin-bottom: 10px;
      display: flex;
    }
    .detail-label {
      font-weight: 600;
      color: #dc2626;
      min-width: 120px;
    }
    .detail-value {
      color: #334155;
      flex: 1;
    }
    .company-brand {
      margin-top: 25px;
      padding-top: 20px;
      border-top: 1px solid #e2e8f0;
      font-size: 12px;
      color: #64748b;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="error-icon">✗</div>
    <h1 class="title">${layerDisplay} Rejection</h1>
    <div class="subtitle">Please provide a reason for rejection</div>
    ${docTypeBadge}
    
    <div class="details">
      <div class="detail-item">
        <span class="detail-label">Project:</span>
        <span class="detail-value">${project || "N/A"}</span>
      </div>
      <div class="detail-item">
        <span class="detail-label">Requester:</span>
        <span class="detail-value">${name || "N/A"}</span>
      </div>
      <div class="detail-item">
        <span class="detail-label">Email:</span>
        <span class="detail-value">${email || "N/A"}</span>
      </div>
      <div class="detail-item">
        <span class="detail-label">Layer:</span>
        <span class="detail-value">${layerDisplay}</span>
      </div>
      <div class="detail-item">
        <span class="detail-label">Date:</span>
        <span class="detail-value">${getGMT7Time()}</span>
      </div>
    </div>

    <div class="rejection-form">
      <div class="form-group">
        <label class="form-label" for="rejectionNote">Rejection Reason (Required):</label>
        <textarea id="rejectionNote" class="form-textarea" 
                  placeholder="Please explain why this request is being rejected. This note will be visible to the requester and previous approvers." 
                  required></textarea>
      </div>
    </div>
    
    <div style="text-align: center;">
      <button type="button" id="submitBtn" class="button" onclick="submitRejectionDirect()">Submit Rejection</button>
    </div>
    
    <div id="loading" style="display: none; color: #dc2626; font-weight: 600; margin-top: 15px;">
      Processing your rejection... Please wait.
    </div>
    
    <div class="company-brand">
      <strong>Atreus Global</strong> • Multi-Layer Approval System
    </div>
  </div>

  <script>
    // PAKE METODE DIRECT REDIRECT - SAMA KAYAK APPROVAL
    function submitRejectionDirect() {
      console.log('Submit button clicked');
      
      var rejectionNote = document.getElementById('rejectionNote').value;
      
      if (!rejectionNote.trim()) {
        alert('Please provide a rejection reason.');
        return;
      }
      
      // Show loading
      document.getElementById('submitBtn').disabled = true;
      document.getElementById('loading').style.display = 'block';
      
      // Build URL dengan rejection note
      var baseUrl = "${WEB_APP_URL}";
      var params = {
        action: "submit_rejection",
        name: "${encodeURIComponent(name || "")}",
        email: "${encodeURIComponent(email || "")}",
        project: "${encodeURIComponent(project || "")}",
        docType: "${encodeURIComponent(documentType || "")}",
        attachment: "${encodeURIComponent(attachment || "")}",
        layer: "${layer}",
        code: "${code}",
        rejectionNote: encodeURIComponent(rejectionNote),
        timestamp: new Date().getTime()
      };
      
      var queryString = Object.keys(params)
        .map(key => key + '=' + params[key])
        .join('&');
      
      var finalUrl = baseUrl + "?" + queryString;
      
      console.log('Redirecting to:', finalUrl);
      
      // Direct redirect - SAMA KAYAK APPROVAL LINK
      window.location.href = finalUrl;
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

// Helper function untuk generate rejection submit link (optional, tapi ga dipake)
function generateRejectionSubmitLink(
  name,
  email,
  description,
  documentType,
  attachment,
  layer,
  code,
  rejectionNote
) {
  var timestamp = new Date().getTime();

  var params = {
    action: "submit_rejection",
    name: encodeURIComponent(name || ""),
    email: encodeURIComponent(email || ""),
    project: encodeURIComponent(description || ""),
    docType: encodeURIComponent(documentType || ""),
    attachment: encodeURIComponent(attachment || ""),
    layer: layer,
    code: code,
    rejectionNote: encodeURIComponent(rejectionNote || ""),
    timestamp: timestamp,
  };

  var queryString = Object.keys(params)
    .map((key) => key + "=" + params[key])
    .join("&");

  return WEB_APP_URL + "?" + queryString;
}

function createRejectionSuccessPage(
  name,
  email,
  project,
  documentType,
  attachment,
  layer,
  rejectionNote
) {
  var layerDisplay = getLayerDisplayName(layer);

  var html =
    '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%;margin:0 auto}.success-icon{font-size:64px;margin-bottom:20px;color:#dc2626}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#dc2626}.details{background:#f8fafc;padding:20px;border-radius:10px;margin:20px 0;text-align:left}.detail-item{margin-bottom:10px}</style></head><body><div class="container"><div class="success-icon">✓</div><h1 class="title">Rejection Submitted</h1><p style="color:#dc2626;margin-bottom:25px;">Document has been sent back for revision</p><div class="details"><div class="detail-item"><strong>Project:</strong> ' +
    (project || "N/A") +
    '</div><div class="detail-item"><strong>Layer:</strong> ' +
    layerDisplay +
    '</div><div class="detail-item"><strong>Date:</strong> ' +
    getGMT7Time() +
    '</div></div><p style="font-size:13px;color:#64748b;margin-top:20px;">Notification has been sent. You can close this page.</p></div><script>setTimeout(function(){window.close()},5000)</script></body></html>';

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function createErrorPage(errorMessage) {
  var html =
    '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%;margin:0 auto}.error-icon{font-size:64px;margin-bottom:20px;color:#dc2626}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#dc2626}.error-box{background:#fef2f2;padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid #dc2626;text-align:left}</style></head><body><div class="container"><div class="error-icon">✗</div><h1 class="title">Error</h1><div class="error-box">' +
    errorMessage +
    '</div><p style="font-size:13px;color:#64748b;margin-top:20px;">Please contact support if this issue persists.</p></div></body></html>';

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function showRejectionSuccessPage(params) {
  var project = params.project ? decodeURIComponent(params.project) : "";
  var layer = params.layer ? decodeURIComponent(params.layer) : "";

  return createRejectionSuccessPage("", "", project, "", "", layer, "");
}

// ============================================
// PART 4: UI FUNCTIONS & MENU (FINAL)
// Multi-Layer Reusable Approval System
// ============================================

// ============================================
// MENU CREATION
// ============================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🔄 Multi-Layer Approval")
    .addItem("📤 Send Approvals", "sendMultiLayerApproval")
    .addItem("📥 Resubmit After Revision", "resubmitAfterRevision")
    .addSeparator()
    .addItem(
      "➡️ Force Send Level Two (After Level One)",
      "sendNextApprovalAfterLevelOne"
    )
    .addItem(
      "➡️ Force Send Level Three (After Level Two)",
      "sendNextApprovalAfterLevelTwo"
    )
    .addItem("➡️ Force Send All Pending Next Layers", "forceSendNextLayers")
    .addSeparator()
    .addItem("📊 View Approval Pipeline", "showApprovalPipeline")
    .addItem("✅ Check Attachment Validation", "validateAllAttachments")
    .addItem("🔄 Reset Selected Rows", "resetMultiLayerRows")
    .addSeparator()
    .addItem("🧪 Test Complete Flow", "testCompleteFlow")
    .addItem("🧪 Test Rejection Flow", "testRejectionFlow")
    .addItem("🔍 Debug Current Row", "debugCurrentRow")
    .addSeparator()
    .addItem("⚙️ Manual Approve - Level One", "manualApproveLevelOne")
    .addItem("⚙️ Manual Approve - Level Two", "manualApproveLevelTwo")
    .addItem("⚙️ Manual Approve - Level Three", "manualApproveLevelThree")
    .addSeparator()
    .addItem("📝 View Recent Logs", "viewLogs")
    .addItem("ℹ️ About System", "showAbout")
    .addToUi();
}

// ============================================
// PIPELINE DASHBOARD
// ============================================

function showApprovalPipeline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:O" + sheet.getLastRow()).getValues();

  var pipeline = {
    PENDING_LEVEL_ONE: [],
    PENDING_LEVEL_TWO: [],
    PENDING_LEVEL_THREE: [],
    EDITING_REQUESTER: [],
    EDITING_LEVEL_ONE: [],
    EDITING_LEVEL_TWO: [],
    COMPLETED: [],
    INVALID_ATTACHMENT: [],
  };

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var levelOne = row[7]; // Column H
    var levelTwo = row[8]; // Column I
    var levelThree = row[9]; // Column J
    var currentEditor = row[13]; // Column N
    var overallStatus = row[14]; // Column O
    var attachment = row[4]; // Column E
    var documentType = row[3]; // Column D

    var itemName = row[0] + " - " + row[2];

    // Check attachment validity
    if (attachment && attachment !== "") {
      var validation = validateGoogleDriveAttachmentWithType(
        attachment,
        documentType
      );
      if (!validation.valid) {
        pipeline.INVALID_ATTACHMENT.push(itemName);
        continue;
      }
    }

    // Categorize based on status
    if (overallStatus === "EDITING") {
      if (currentEditor === "REQUESTER") {
        pipeline.EDITING_REQUESTER.push(itemName);
      } else if (currentEditor === "LEVEL_ONE") {
        pipeline.EDITING_LEVEL_ONE.push(itemName);
      } else if (currentEditor === "LEVEL_TWO") {
        pipeline.EDITING_LEVEL_TWO.push(itemName);
      }
    } else if (overallStatus === "COMPLETED") {
      pipeline.COMPLETED.push(itemName);
    } else {
      // Check approval stages
      if (!levelOne || levelOne === "PENDING" || levelOne === "RESUBMIT") {
        pipeline.PENDING_LEVEL_ONE.push(itemName);
      } else if (
        levelOne === "APPROVED" &&
        (!levelTwo || levelTwo === "PENDING" || levelTwo === "RESUBMIT")
      ) {
        pipeline.PENDING_LEVEL_TWO.push(itemName);
      } else if (
        levelTwo === "APPROVED" &&
        (!levelThree || levelThree === "PENDING" || levelThree === "RESUBMIT")
      ) {
        pipeline.PENDING_LEVEL_THREE.push(itemName);
      }
    }
  }

  var message = "📊 MULTI-LAYER APPROVAL PIPELINE\n\n";
  message += "⏳ PENDING APPROVALS:\n";
  message += "  • Level One: " + pipeline.PENDING_LEVEL_ONE.length + "\n";
  message += "  • Level Two: " + pipeline.PENDING_LEVEL_TWO.length + "\n";
  message += "  • Level Three: " + pipeline.PENDING_LEVEL_THREE.length + "\n\n";

  message += "✏️ CURRENTLY EDITING:\n";
  message += "  • Requester: " + pipeline.EDITING_REQUESTER.length + "\n";
  message += "  • Level One: " + pipeline.EDITING_LEVEL_ONE.length + "\n";
  message += "  • Level Two: " + pipeline.EDITING_LEVEL_TWO.length + "\n\n";

  message += "✅ Completed: " + pipeline.COMPLETED.length + "\n";
  message +=
    "❌ Invalid Attachment: " + pipeline.INVALID_ATTACHMENT.length + "\n\n";

  if (pipeline.INVALID_ATTACHMENT.length > 0) {
    message += "⚠️ INVALID ATTACHMENTS:\n";
    pipeline.INVALID_ATTACHMENT.forEach(function (item) {
      message += "  • " + item + "\n";
    });
  }

  SpreadsheetApp.getUi().alert(
    "Approval Pipeline Dashboard",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================
// VALIDATION FUNCTIONS
// ============================================

function validateAllAttachments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:O" + sheet.getLastRow()).getValues();

  var validationResults = {
    valid: 0,
    invalid: 0,
    empty: 0,
    sharedDrive: 0,
    myDrive: 0,
    details: [],
  };

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var attachment = row[4]; // Column E
    var documentType = row[3]; // Column D
    var validation = validateGoogleDriveAttachmentWithType(
      attachment,
      documentType
    );

    validationResults.details.push({
      row: i + 2,
      project: row[2],
      documentType: documentType,
      attachment: attachment,
      validation: validation,
    });

    if (!attachment || attachment === "") {
      validationResults.empty++;
    } else if (validation.valid) {
      validationResults.valid++;
      if (validation.isSharedDrive) {
        validationResults.sharedDrive++;
      } else {
        validationResults.myDrive++;
      }
    } else {
      validationResults.invalid++;
    }
  }

  var message = "📎 ATTACHMENT VALIDATION REPORT\n\n";
  message += "✅ Valid: " + validationResults.valid + "\n";
  message += "  • Shared Drive: " + validationResults.sharedDrive + "\n";
  message += "  • My Drive: " + validationResults.myDrive + "\n";
  message += "❌ Invalid: " + validationResults.invalid + "\n";
  message += "⚪ Empty: " + validationResults.empty + "\n\n";

  if (validationResults.invalid > 0) {
    message += "⚠️ INVALID ATTACHMENTS FOUND:\n";
    validationResults.details.forEach(function (detail) {
      if (!detail.validation.valid && detail.attachment) {
        message +=
          "• Row " +
          detail.row +
          ": " +
          detail.project +
          "\n  → " +
          detail.validation.message +
          "\n";
      }
    });
  }

  SpreadsheetApp.getUi().alert(
    "Attachment Validation",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================
// RESET & UTILITY FUNCTIONS
// ============================================

function resetMultiLayerRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = sheet.getActiveRange();
  var startRow = activeRange.getRow();

  if (startRow < 2) {
    SpreadsheetApp.getUi().alert(
      "Invalid Selection",
      "Please select rows starting from row 2.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var confirmation = SpreadsheetApp.getUi().alert(
    "Confirm Reset",
    "Are you sure you want to reset " +
      activeRange.getNumRows() +
      " row(s)?\n\nThis will:\n• Clear all approval statuses\n• Reset to PENDING state\n• Clear logs and notes\n• Set status to ACTIVE",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (confirmation !== SpreadsheetApp.getUi().Button.YES) {
    return;
  }

  for (var i = startRow; i < startRow + activeRange.getNumRows(); i++) {
    if (i >= 2) {
      // Reset checkboxes and status
      sheet.getRange(i, 6).setValue(false); // Column F - Send Checkbox
      sheet.getRange(i, 7).setValue(""); // Column G - Log
      sheet.getRange(i, 7).setNote(""); // Clear note

      // Reset all layer statuses
      sheet.getRange(i, 8).setValue("PENDING"); // Column H - Level One Status
      sheet.getRange(i, 8).setBackground(null);
      sheet.getRange(i, 8).setNote("");

      sheet.getRange(i, 9).setValue("PENDING"); // Column I - Level Two Status
      sheet.getRange(i, 9).setBackground(null);
      sheet.getRange(i, 9).setNote("");

      sheet.getRange(i, 10).setValue("PENDING"); // Column J - Level Three Status
      sheet.getRange(i, 10).setBackground(null);
      sheet.getRange(i, 10).setNote("");

      // Clear current editor
      sheet.getRange(i, 14).setValue(""); // Column N - Current Editor

      // Reset overall status
      sheet.getRange(i, 15).setValue("ACTIVE"); // Column O - Overall Status
      sheet.getRange(i, 15).setBackground(null);
    }
  }

  SpreadsheetApp.getUi().alert(
    "Reset Complete",
    activeRange.getNumRows() + " row(s) have been reset to PENDING state!",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function forceSendNextLayers() {
  try {
    Logger.log("Force sending all pending next layers...");
    sendNextApprovalAfterLevelOne();
    Utilities.sleep(2000);
    sendNextApprovalAfterLevelTwo();
    SpreadsheetApp.getUi().alert(
      "Force Send Complete",
      "✅ Checked and sent all pending next layer approvals.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log("❌ Error in forceSendNextLayers: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "Error",
      "Force send error: " + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================
// MANUAL APPROVAL FUNCTIONS
// ============================================

function manualApproveRow(rowNumber, layer) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rowData = sheet.getRange(rowNumber, 1, 1, 15).getValues()[0];

    var name = rowData[0];
    var email = rowData[1];
    var project = rowData[2];
    var code = "manual_" + new Date().getTime();

    var result = updateMultiLayerApprovalStatus(
      name,
      email,
      project,
      layer,
      code
    );

    if (result) {
      SpreadsheetApp.getUi().alert(
        "Success",
        "✅ Manual approval completed for:\n\n" +
          name +
          "\nProject: " +
          project +
          "\nLayer: " +
          getLayerDisplayName(layer),
        SpreadsheetApp.getUi().ButtonSet.OK
      );

      // Auto-trigger next layer
      if (layer === "LEVEL_ONE") {
        Utilities.sleep(2000);
        sendNextApprovalAfterLevelOne();
      } else if (layer === "LEVEL_TWO") {
        Utilities.sleep(2000);
        sendNextApprovalAfterLevelTwo();
      }
    } else {
      SpreadsheetApp.getUi().alert(
        "Failed",
        "❌ Manual approval failed for " + name + ". Check logs for details.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log("❌ Error in manualApproveRow: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "Error",
      "Manual approval error: " + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function manualApproveLevelOne() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert(
      "Invalid Row",
      "Please select a row starting from row 2.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  manualApproveRow(row, "LEVEL_ONE");
}

function manualApproveLevelTwo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert(
      "Invalid Row",
      "Please select a row starting from row 2.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  manualApproveRow(row, "LEVEL_TWO");
}

function manualApproveLevelThree() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert(
      "Invalid Row",
      "Please select a row starting from row 2.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  manualApproveRow(row, "LEVEL_THREE");
}

// ============================================
// DEBUG & TEST FUNCTIONS
// ============================================

function debugCurrentRow() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = sheet.getActiveRange().getRow();

    if (row < 2) {
      SpreadsheetApp.getUi().alert(
        "Invalid Row",
        "Please select a row starting from row 2.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    var rowData = sheet.getRange(row, 1, 1, 15).getValues()[0];

    var debugInfo = "🔍 DEBUG INFO - Row " + row + "\n\n";
    debugInfo += "👤 REQUESTER INFO:\n";
    debugInfo += "  • Name: " + (rowData[0] || "Empty") + "\n";
    debugInfo += "  • Email: " + (rowData[1] || "Empty") + "\n";
    debugInfo += "  • Project: " + (rowData[2] || "Empty") + "\n";
    debugInfo += "  • Doc Type: " + (rowData[3] || "Empty") + "\n\n";

    debugInfo += "📎 ATTACHMENT:\n";
    debugInfo += "  • URL: " + (rowData[4] || "Empty") + "\n\n";

    debugInfo += "✅ APPROVAL STATUS:\n";
    debugInfo += "  • Send Checkbox: " + rowData[5] + "\n";
    debugInfo += "  • Level One: " + (rowData[7] || "PENDING") + "\n";
    debugInfo += "  • Level Two: " + (rowData[8] || "PENDING") + "\n";
    debugInfo += "  • Level Three: " + (rowData[9] || "PENDING") + "\n\n";

    debugInfo += "📧 APPROVER EMAILS:\n";
    debugInfo += "  • Level One: " + (rowData[10] || "Empty") + "\n";
    debugInfo += "  • Level Two: " + (rowData[11] || "Empty") + "\n";
    debugInfo += "  • Level Three: " + (rowData[12] || "Empty") + "\n\n";

    debugInfo += "🔄 TRACKING:\n";
    debugInfo += "  • Current Editor: " + (rowData[13] || "None") + "\n";
    debugInfo += "  • Overall Status: " + (rowData[14] || "ACTIVE") + "\n";

    // Validate attachment
    if (rowData[4]) {
      var validation = validateGoogleDriveAttachmentWithType(
        rowData[4],
        rowData[3]
      );
      debugInfo += "\n📋 ATTACHMENT VALIDATION:\n";
      debugInfo +=
        "  • Valid: " + (validation.valid ? "✅ Yes" : "❌ No") + "\n";
      debugInfo += "  • Message: " + validation.message + "\n";
      if (validation.valid) {
        debugInfo += "  • File: " + validation.name + "\n";
        debugInfo += "  • Type: " + validation.type + "\n";
        debugInfo += "  • Size: " + validation.sizeFormatted + "\n";
        debugInfo += "  • Location: " + validation.driveType + "\n";
      }
    }

    SpreadsheetApp.getUi().alert(
      "Debug Information",
      debugInfo,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log("❌ Error in debugCurrentRow: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "Debug Error",
      "Error: " + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function testCompleteFlow() {
  Logger.log("🧪 Testing Complete Approval Flow...");

  var testName = "Test User";
  var testEmail = "test@atreusg.com";
  var testProject = "Test Project Flow";
  var testLayer = "LEVEL_ONE";
  var testCode = "test_" + new Date().getTime();

  var result = updateMultiLayerApprovalStatus(
    testName,
    testEmail,
    testProject,
    testLayer,
    testCode
  );

  if (result) {
    Logger.log("✅ Test PASSED - Approval workflow is working!");
    SpreadsheetApp.getUi().alert(
      "Test Result",
      "✅ Test PASSED\n\nApproval workflow is working correctly!",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    Logger.log("❌ Test FAILED - Check logs for details");
    SpreadsheetApp.getUi().alert(
      "Test Result",
      "❌ Test FAILED\n\nCheck logs (Extensions > Apps Script > Executions) for details.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function testRejectionFlow() {
  Logger.log("🧪 Testing Rejection & Send Back Flow...");

  var testData = [
    { layer: "LEVEL_ONE", name: "Test User 1", project: "Test Reject L1" },
    { layer: "LEVEL_TWO", name: "Test User 2", project: "Test Reject L2" },
    { layer: "LEVEL_THREE", name: "Test User 3", project: "Test Reject L3" },
  ];

  var results = [];

  testData.forEach(function (test) {
    Logger.log("Testing " + test.layer + " rejection...");

    var result = updateMultiLayerRejectionStatus(
      test.name,
      "test@atreusg.com",
      test.project,
      test.layer,
      "test_" + new Date().getTime(),
      "Test rejection reason for " + test.layer
    );

    if (result) {
      Logger.log("✅ " + test.layer + " rejection: PASSED");
      results.push("✅ " + test.layer + ": PASSED");
    } else {
      Logger.log("❌ " + test.layer + " rejection: FAILED");
      results.push("❌ " + test.layer + ": FAILED");
    }
  });

  var message =
    "🧪 REJECTION FLOW TEST RESULTS\n\n" +
    results.join("\n") +
    "\n\nCheck the spreadsheet for:\n• Status changed to REJECTED\n• Previous layer set to RESUBMIT\n• Current Editor updated\n• Overall Status set to EDITING";

  SpreadsheetApp.getUi().alert(
    "Test Results",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function viewLogs() {
  var message = "📝 VIEW EXECUTION LOGS\n\n";
  message += "To view detailed logs:\n\n";
  message += "1. Go to Extensions > Apps Script\n";
  message += "2. Click 'Executions' on the left sidebar\n";
  message += "3. View recent runs and their logs\n\n";
  message += "You can also:\n";
  message += "• Filter by status (Success/Error)\n";
  message += "• Search for specific functions\n";
  message += "• Export logs for analysis";

  SpreadsheetApp.getUi().alert(
    "View Logs",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showAbout() {
  var message = "🔄 MULTI-LAYER REUSABLE APPROVAL SYSTEM\n\n";
  message += "Version: 2.0 (Reusable)\n";
  message +=
    "Updated: " +
    Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy") +
    "\n\n";
  message += "FEATURES:\n";
  message += "✅ 3-layer approval workflow\n";
  message += "✅ Reusable rows (no duplication)\n";
  message += "✅ Send back to previous layer on rejection\n";
  message += "✅ Resubmit capability after editing\n";
  message += "✅ Google Drive attachment validation\n";
  message += "✅ Email notifications with HTML templates\n";
  message += "✅ Real-time status tracking\n\n";
  message += "WORKFLOW:\n";
  message += "• Level One reject → Requester edits\n";
  message += "• Level Two reject → Level One edits\n";
  message += "• Level Three reject → Level Two edits\n\n";
  message += "© 2024 Atreus Global";

  SpreadsheetApp.getUi().alert(
    "About System",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function resubmitAfterRevision() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = sheet.getActiveRange().getRow();

    if (row < 2) {
      SpreadsheetApp.getUi().alert(
        "Error",
        "Please select a row that needs to be resubmitted.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get row data
    var currentEditor = sheet.getRange(row, 14).getValue(); // Column N - Current Editor
    var attachment = sheet.getRange(row, 5).getValue(); // Column E - Attachment
    var documentType = sheet.getRange(row, 4).getValue(); // Column D - Document Type
    var name = sheet.getRange(row, 1).getValue(); // Column A - Name
    var email = sheet.getRange(row, 2).getValue(); // Column B - Email
    var description = sheet.getRange(row, 3).getValue(); // Column C - Description

    // Validate if user can resubmit
    if (
      currentEditor !== "REQUESTER" &&
      currentEditor !== "LEVEL_ONE" &&
      currentEditor !== "LEVEL_TWO"
    ) {
      SpreadsheetApp.getUi().alert(
        "Error",
        "This row is not currently in edit mode.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Validate attachment
    if (!attachment) {
      SpreadsheetApp.getUi().alert(
        "Error",
        "Please update the Google Drive link in Column E (Attachment) before resubmitting.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    var validation = validateGoogleDriveAttachmentWithType(
      attachment,
      documentType
    );
    if (!validation.valid) {
      SpreadsheetApp.getUi().alert(
        "Error",
        "Invalid Google Drive link. Please check the attachment:\n\n" +
          validation.message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Confirmation
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      "Confirm Resubmit",
      "Are you sure you want to resubmit this document for approval?\n\nThis will:\n• Validate the new attachment\n• Reset the current level status\n• Send for next level review",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Process resubmit based on current editor
    if (currentEditor === "REQUESTER") {
      // Requester resubmitting after Level One rejection
      sheet.getRange(row, 8).setValue("PENDING"); // Reset Level One status
      sheet.getRange(row, 8).setBackground(null);
      sheet.getRange(row, 14).setValue(""); // Clear current editor
      sheet.getRange(row, 15).setValue("PROCESSING"); // Set overall status
      sheet.getRange(row, 15).setBackground("#FFF2CC");

      // Trigger approval process
      sendMultiLayerApproval();
    } else if (currentEditor === "LEVEL_ONE") {
      // Level One resubmitting after Level Two rejection
      sheet.getRange(row, 9).setValue("PENDING"); // Reset Level Two status
      sheet.getRange(row, 9).setBackground(null);
      sheet.getRange(row, 8).setValue("APPROVED"); // Keep Level One as approved
      sheet.getRange(row, 8).setBackground("#90EE90");
      sheet.getRange(row, 14).setValue(""); // Clear current editor
      sheet.getRange(row, 15).setValue("PROCESSING"); // Set overall status
      sheet.getRange(row, 15).setBackground("#FFF2CC");

      // Trigger next level approval
      sendNextApprovalAfterLevelOne();
    } else if (currentEditor === "LEVEL_TWO") {
      // Level Two resubmitting after Level Three rejection
      sheet.getRange(row, 10).setValue("PENDING"); // Reset Level Three status
      sheet.getRange(row, 10).setBackground(null);
      sheet.getRange(row, 9).setValue("APPROVED"); // Keep Level Two as approved
      sheet.getRange(row, 9).setBackground("#90EE90");
      sheet.getRange(row, 14).setValue(""); // Clear current editor
      sheet.getRange(row, 15).setValue("PROCESSING"); // Set overall status
      sheet.getRange(row, 15).setBackground("#FFF2CC");

      // Trigger next level approval
      sendNextApprovalAfterLevelTwo();
    }

    // Add resubmission note
    var currentTime = getGMT7Time();
    var existingNote = sheet.getRange(row, 7).getNote() || "";
    var resubmitNote =
      "Resubmitted by " +
      currentEditor +
      " on " +
      currentTime +
      "\n" +
      existingNote;
    sheet.getRange(row, 7).setNote(resubmitNote);

    SpreadsheetApp.getUi().alert(
      "Success",
      "✅ Document has been resubmitted successfully!\n\nThe approval request will be sent to the next level automatically.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log("❌ Error in resubmitAfterRevision: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "Error",
      "Failed to resubmit: " +
        error.message +
        "\n\nPlease try again or contact support.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function getCellValue(sheet, row, column) {
  try {
    return sheet.getRange(row, column).getValue();
  } catch (e) {
    return "";
  }
}

function getApprovalCodeFromLink(approvalLink) {
  var match = approvalLink.match(/code=([^&]+)/);
  return match ? match[1] : "N/A";
}

// ============================================
// SCHEDULED TRIGGER FUNCTIONS (OPTIONAL)
// ============================================

/**
 * Set up time-based triggers untuk auto-send approval
 * Jalankan function ini sekali untuk setup trigger
 */
function setupAutomationTriggers() {
  // Delete existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  // Create new trigger: Check every hour for pending approvals
  ScriptApp.newTrigger("autoCheckPendingApprovals")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("✅ Automation triggers set up successfully");
  SpreadsheetApp.getUi().alert(
    "Triggers Setup",
    "✅ Automation triggers have been set up!\n\nThe system will now automatically check for pending approvals every hour.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function autoCheckPendingApprovals() {
  Logger.log("⏰ Auto-check: Running scheduled approval check...");

  try {
    // Send next layer approvals automatically
    sendNextApprovalAfterLevelOne();
    Utilities.sleep(2000);
    sendNextApprovalAfterLevelTwo();

    Logger.log("✅ Auto-check completed successfully");
  } catch (error) {
    Logger.log("❌ Auto-check error: " + error.toString());
    sendAdminNotification(
      "Auto-check failed: " + error.message,
      "Approval System Error"
    );
  }
}

/**
 * Remove all automation triggers
 */
function removeAutomationTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  Logger.log("✅ All automation triggers removed");
  SpreadsheetApp.getUi().alert(
    "Triggers Removed",
    "✅ All automation triggers have been removed.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================
// END OF PART 4
// ============================================
