// ============================================
// PART 1: COMPLETE FIXED VERSION - MAIN FUNCTIONS
// Adjusted for your sheet structure
// ============================================

var DOCUMENT_TYPE_FOLDERS = {
  "ICC": "1JFJPfirJuCvZuSEKe6KwRXuupRI0hyyb",
  "Quotation": "1QVTM_oTSQow9N0e1jNIAc-ADK0qvTwtz",
  "Proposal": "1QVTM_oTSQow9N0e1jNIAc-ADK0qvTwtz"
};

function sendMultiLayerApproval() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("No data to process!");
      return;
    }
    
    // FIXED: Correct range based on your sheet (A:N columns)
    var dataRange = sheet.getRange("A2:N" + lastRow);
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
      var currentStatus = row[13]; // Column N - Status
      
      if (!name || !email) continue;
      
      var isChecked = (sendStatus === true || sendStatus === "TRUE" || sendStatus === "true");
      var isActive = (currentStatus === "ACTIVE" || currentStatus === "" || !currentStatus);
      
      if (isChecked && isActive) {
        Logger.log("Processing: " + name + " - " + description);
        
        var validationResult = validateGoogleDriveAttachmentWithType(attachment, documentType);
        
        // FIXED: Correct column indexes based on your sheet
        var firstLayerStatus = row[7]; // Column H - FirstLayer
        var secondLayerStatus = row[8]; // Column I - Second Layer  
        var thirdLayerStatus = row[9]; // Column J - ThirdLayer

        var firstLayerEmail = row[10]; // Column K - FirstLayerEmail
        var secondLayerEmail = row[11]; // Column L - SecondLayer Email
        var thirdLayerEmail = row[12]; // Column M - ThirdLayerEmall
        
        var nextApproval = getNextApprovalLayerAndEmail(
          firstLayerStatus, secondLayerStatus, thirdLayerStatus,
          firstLayerEmail, secondLayerEmail, thirdLayerEmail
        );
        
        Logger.log("Next approval: " + nextApproval.layer + " -> " + nextApproval.email);
        
        if (nextApproval.layer !== "COMPLETED" && nextApproval.email) {
          var approvalLink = generateMultiLayerApprovalLink(name, email, description, documentType, attachment, nextApproval.layer);
          var emailSent = sendMultiLayerEmail(nextApproval.email, description, documentType, attachment, approvalLink, nextApproval.layer, validationResult);
          
          if (emailSent) {
            // Store approval link in notes (Column G - Log)
            sheet.getRange(i + 2, 7).setNote("APPROVAL_LINK_" + nextApproval.layer + ": " + approvalLink);
            sheet.getRange(i + 2, 14).setValue("PROCESSING");
            sheet.getRange(i + 2, 14).setBackground("#FFF2CC");
            
            results.push({
              name: name,
              project: description,
              layer: nextApproval.layer,
              status: "Email Sent to: " + nextApproval.email
            });
            
            processedCount++;
            Logger.log(" Approval email sent for: " + name);
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
    SpreadsheetApp.getUi().alert("System Error", "Error: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getNextApprovalLayerAndEmail(firstLayer, secondLayer, thirdLayer, firstLayerEmail, secondLayerEmail, thirdLayerEmail) {
  Logger.log("Checking approval flow:");
  Logger.log("First Layer: " + firstLayer + ", Second Layer: " + secondLayer + ", Third Layer: " + thirdLayer);

  if (!firstLayer || firstLayer === "" || firstLayer === "PENDING") {
    Logger.log("Next: First Layer");
    return { layer: "FIRST_LAYER", email: firstLayerEmail };
  }
  if (firstLayer === "APPROVED" && (!secondLayer || secondLayer === "" || secondLayer === "PENDING")) {
    Logger.log("Next: Second Layer");
    return { layer: "SECOND_LAYER", email: secondLayerEmail };
  }
  if (secondLayer === "APPROVED" && (!thirdLayer || thirdLayer === "" || thirdLayer === "PENDING")) {
    Logger.log("Next: Third Layer");
    return { layer: "THIRD_LAYER", email: thirdLayerEmail };
  }
  Logger.log("Next: COMPLETED");
  return { layer: "COMPLETED", email: "" };
}

function generateMultiLayerApprovalLink(name, email, description, documentType, attachment, layer, action) {
  // FIXED: Use your actual web app URL
  var webAppUrl = "https://script.google.com/macros/s/AKfycbw-7cNeUsM82jvqztSJ4GpSI4Nzdg2Fq-yS-2kAGDkUyKF-uvX3MDVsvQXuqNKcxBd-/exec";

  var timestamp = new Date().getTime().toString(36);
  var nameCode = name ? name.substring(0, 3).toLowerCase() : "usr";
  var projectCode = description ? description.replace(/[^a-zA-Z0-9]/g, "").substring(0, 3).toLowerCase() : "prj";
  var uniqueCode = nameCode + projectCode + timestamp;

  // FIXED: Ensure all parameters are properly encoded and have fallbacks
  var params = {
    action: action || "approve",
    name: encodeURIComponent(name || ""),
    email: encodeURIComponent(email || ""),
    project: encodeURIComponent(description || ""),
    docType: encodeURIComponent(documentType || ""),
    attachment: encodeURIComponent(attachment || ""),
    layer: layer,
    code: uniqueCode,
    timestamp: new Date().getTime()
  };

  var queryString = Object.keys(params)
    .map(key => key + '=' + params[key])
    .join('&');

  return webAppUrl + "?" + queryString;
}

// ============================================
// FIXED: VALIDATION FUNCTIONS
// ============================================

function validateGoogleDriveAttachmentWithType(attachmentUrl, documentType) {
  try {
    if (!attachmentUrl || attachmentUrl === "") {
      return {
        valid: false,
        message: "No attachment provided",
        type: "EMPTY",
        isSharedDrive: false
      };
    }
    
    if (!attachmentUrl.includes('drive.google.com')) {
      return {
        valid: false,
        message: "Not a Google Drive link",
        type: "INVALID_URL",
        isSharedDrive: false
      };
    }
    
    var fileId = extractFileIdFromUrl(attachmentUrl);
    if (!fileId) {
      return {
        valid: false,
        message: "Invalid Google Drive URL format",
        type: "INVALID_FORMAT",
        isSharedDrive: false
      };
    }
    
    try {
      var file = DriveApp.getFileById(fileId);
      
      if (file.getMimeType() === 'application/vnd.google-apps.folder') {
        return {
          valid: false,
          message: "Google Drive folder detected - please use file link",
          type: "FOLDER_NOT_SUPPORTED",
          isSharedDrive: false
        };
      }
      
      // FIXED: Check if file is in Shared Drive (safe version)
      var isSharedDrive = false;
      try {
        var driveFile = Drive.Files.get(fileId, {
          supportsAllDrives: true,
          fields: 'driveId'
        });
        isSharedDrive = driveFile.driveId != null;
      } catch (e) {
        Logger.log("Note: Could not check Shared Drive status: " + e.toString());
        isSharedDrive = false;
      }
      
      // FIXED: Safe owner retrieval
      var ownerEmail = "Unknown";
      try {
        var owner = file.getOwner();
        if (owner && owner.getEmail) {
          ownerEmail = owner.getEmail();
        }
      } catch (ownerError) {
        ownerEmail = isSharedDrive ? "Shared Drive" : "Unknown";
      }
      
      // Document type info
      var folderInfo = "";
      if (documentType) {
        folderInfo = "Document type: " + documentType;
      }
      
      return {
        valid: true,
        message: "Valid Google Drive file",
        type: file.getMimeType(),
        name: file.getName(),
        size: file.getSize(),
        sizeFormatted: formatFileSize(file.getSize()),
        url: file.getUrl(),
        owner: ownerEmail,
        lastUpdated: file.getLastUpdated(),
        folderWarning: null,
        folderInfo: folderInfo,
        isSharedDrive: isSharedDrive,
        driveType: isSharedDrive ? "Shared Drive" : "My Drive"
      };
      
    } catch (e) {
      Logger.log("Error accessing file: " + e.toString());
      return {
        valid: false,
        message: "File not found or no access permission",
        type: "NO_ACCESS",
        isSharedDrive: false
      };
    }
    
  } catch (error) {
    Logger.log("Validation error: " + error.toString());
    return {
      valid: false,
      message: "Validation error: " + error.message,
      type: "VALIDATION_ERROR",
      isSharedDrive: false
    };
  }
}

function extractFileIdFromUrl(url) {
  var patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/open\?id=([a-zA-Z0-9_-]+)/
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
  if (bytes === 0) return '0 Bytes';
  var k = 1024;
  var sizes = ['Bytes', 'KB', 'MB', 'GB'];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// ============================================
// EMAIL FUNCTIONS
// ============================================

function sendMultiLayerEmail(recipientEmail, description, documentType, attachment, approvalLink, layer, validationResult) {
  try {
    if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
      Logger.log("Invalid email: " + recipientEmail);
      return false;
    }
    
    var layerDisplayNames = {
      "FIRST_LAYER": "First Layer",
      "SECOND_LAYER": "Second Layer",
      "THIRD_LAYER": "Third Layer"
    };
    
    var layerDisplay = layerDisplayNames[layer] || layer;
    var companyName = "Atreus Global";
    
    var subject = "Approval Request - " + description + " [" + layerDisplay + " Layer]";
    
    var docTypeBadge = "";
    if (documentType && documentType !== "") {
      var docTypeColors = {
        "ICC": "#3B82F6",
        "Quotation": "#10B981",
        "Proposal": "#F59E0B"
      };
      var badgeColor = docTypeColors[documentType] || "#6B7280";
      docTypeBadge = '<div style="display: inline-block; background: ' + badgeColor + '; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-bottom: 10px;">' + documentType + '</div>';
    }
    
    var attachmentSection = "";
    if (attachment && attachment !== "") {
      if (validationResult.valid) {
        var driveTypeInfo = validationResult.isSharedDrive ? 
          '<p><strong>Location:</strong> <span style="color: #3B82F6;">📁 Shared Drive</span></p>' :
          '<p><strong>Location:</strong> My Drive</p>';
        
        attachmentSection = '<div class="attachment-box" style="background: linear-gradient(135deg, #f0f8ff 0%, #e6f2ff 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #90EE90;"><h3 style="color: #2E8B57;">Attachment (Validated)</h3>' + docTypeBadge + '<p><strong>File:</strong> ' + validationResult.name + '</p><p><strong>Type:</strong> ' + validationResult.type + '</p><p><strong>Size:</strong> ' + validationResult.sizeFormatted + '</p>' + driveTypeInfo + '<p><strong>Status:</strong> <span style="color: #2E8B57;">' + validationResult.message + '</span></p><p><a href="' + attachment + '" target="_blank" style="color: #326BC6;">View Document</a></p></div>';
      } else {
        attachmentSection = '<div class="attachment-box" style="background: linear-gradient(135deg, #fff0f0 0%, #ffe6e6 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #FF6B6B;"><h3 style="color: #DC143C;">Attachment (Validation Failed)</h3>' + docTypeBadge + '<p><strong>Status:</strong> <span style="color: #DC143C;">' + validationResult.message + '</span></p><p><a href="' + attachment + '" target="_blank" style="color: #326BC6;">Check Link</a></p><p style="font-size: 12px; color: #666;"><em>Please ensure the Google Drive link is correct and accessible</em></p></div>';
      }
    }
    
    var progressBar = getProgressBar(layer);
    
    // FIXED: Improved styling with proper spacing and white button text
    var rejectLink = generateMultiLayerApprovalLink("", "", description, documentType, attachment, layer, "reject");
    var htmlBody = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;line-height:1.6;color:#333;background:#f6f9fc;margin:0;padding:0}.container{max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1)}.header{background:linear-gradient(135deg,#326BC6 0%,#183460 100%);padding:30px;text-align:center;color:white}.progress-track{background:#f8f9fa;padding:30px 20px;margin:0}.progress-steps{display:flex;justify-content:center;align-items:center;gap:40px;position:relative}.progress-step{text-align:center;position:relative}.step-number{width:40px;height:40px;border-radius:50%;margin:0 auto 8px;font-weight:600;display:flex;align-items:center;justify-content:center;font-size:16px}.step-active{background:#326BC6;color:white}.step-completed{background:#90EE90;color:#333}.step-pending{background:#e0e0e0;color:#666}.step-label{font-size:12px;font-weight:500}.content{padding:30px}.button{display:inline-block;padding:15px 30px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white !important;text-decoration:none;border-radius:8px;margin:20px 10px;font-weight:600;font-size:16px;border:none;cursor:pointer;transition:all 0.3s ease}.button:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(50,107,198,0.3)}.button-reject{background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white !important;text-decoration:none;border-radius:8px;margin:20px 10px;font-weight:600;font-size:16px;border:none;cursor:pointer;transition:all 0.3s ease}.button-reject:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(220,38,38,0.3)}.button-container{text-align:center;margin:20px 0}.info-box{background:#f8f9fa;padding:15px;border-radius:5px;margin:15px 0;border-left:4px solid #326BC6}.footer{margin-top:30px;padding:20px;background:#f8f9fa;text-align:center;font-size:12px;color:#666}</style></head><body><div class="container"><div class="header"><h1>Multi-Layer Approval Required</h1><p>Current Stage: ' + layerDisplay + ' Approval</p></div><div class="progress-track">' + progressBar + '</div><div class="content"><p>Hello,</p><p>This project requires your approval at the <strong>' + layerDisplay + '</strong> level:</p><div class="info-box"><h3>' + description + '</h3><p><strong>Current Stage:</strong> ' + layerDisplay + ' Approval</p><p><strong>Date:</strong> ' + getGMT7Time() + '</p></div>' + attachmentSection + '<p>Please choose your decision:</p><div class="button-container"><a href="' + approvalLink + '" class="button" style="color: white !important;">APPROVE AS ' + layerDisplay + '</a><a href="' + rejectLink + '" class="button-reject" style="color: white !important;">REJECT WITH NOTE</a></div><p style="text-align: center; font-size: 12px; color: #666;"><em>Approval Code: ' + getApprovalCodeFromLink(approvalLink) + ' | Link expires in 3 days</em></p></div><div class="footer"><p>This is an automated email. Please do not reply.</p><p>© ' + new Date().getFullYear() + ' ' + companyName + '. All rights reserved.</p></div></div></body></html>';
    
    var plainBody = "MULTI-LAYER APPROVAL REQUEST\n\nProject: " + description + "\nDocument Type: " + documentType + "\nCurrent Stage: " + layerDisplay + " Approval\n\n";
    
    if (attachment && attachment !== "") {
      plainBody += "Attachment: " + attachment + "\nAttachment Status: " + validationResult.message + "\n\n";
    }
    
    plainBody += "Approve: " + approvalLink + "\nReject: " + rejectLink + "\n\nThank you.\n\nThis is an automated email from " + companyName + ".";
    
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody
    });
    
    Logger.log(" Multi-layer email sent to: " + recipientEmail + " for layer: " + layerDisplay);
    return true;
    
  } catch (error) {
    Logger.log("Error sending multi-layer email: " + error.toString());
    return false;
  }
}

function getProgressBar(currentLayer) {
  var steps = [
    { number: 1, label: "First Layer", status: currentLayer === "FIRST_LAYER" ? "active" : (["SECOND_LAYER", "THIRD_LAYER", "COMPLETED"].includes(currentLayer) ? "completed" : "pending") },
    { number: 2, label: "Second Layer", status: currentLayer === "SECOND_LAYER" ? "active" : (["THIRD_LAYER", "COMPLETED"].includes(currentLayer) ? "completed" : "pending") },
    { number: 3, label: "Third Layer", status: currentLayer === "THIRD_LAYER" ? "active" : (currentLayer === "COMPLETED" ? "completed" : "pending") }
  ];
  
  var progressHtml = '<div class="progress-steps">';
  
  steps.forEach(function(step) {
    progressHtml += '<div class="progress-step"><div class="step-number step-' + step.status + '">' + step.number + '</div><div class="step-label">' + step.label + '</div></div>';
  });
  
  progressHtml += '</div>';
  return progressHtml;
}

function getApprovalCodeFromLink(approvalLink) {
  var match = approvalLink.match(/code=([^&]+)/);
  return match ? match[1] : "N/A";
}

function getGMT7Time() {
  var now = new Date();
  var timeZone = "Asia/Jakarta";
  var formattedDate = Utilities.formatDate(now, timeZone, "dd/MM/yyyy HH:mm:ss");
  return formattedDate;
}

// ============================================
// PART 2: WEB APP HANDLERS & APPROVAL FLOW
// ============================================

function showMultiLayerSummary(results, processedCount) {
  if (processedCount > 0) {
    var message = "Multi-Layer Approval Summary\n\nTotal processed: " + processedCount + "\n\n";
    
    var layerCount = {};
    results.forEach(function(result) {
      if (!layerCount[result.layer]) layerCount[result.layer] = 0;
      layerCount[result.layer]++;
    });
    
    for (var layer in layerCount) {
      var layerName = getLayerDisplayName(layer);
      message += layerName + " layer: " + layerCount[layer] + " emails\n";
    }
    
    SpreadsheetApp.getUi().alert("Multi-Layer Approval Sent", message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("No Action Needed", "No multi-layer approvals to send.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getLayerDisplayName(layerCode) {
  var layerNames = {
    "FIRST_LAYER": "First Layer",
    "SECOND_LAYER": "Second Layer",
    "THIRD_LAYER": "Third Layer"
  };
  return layerNames[layerCode] || layerCode;
}

function sendAdminNotification(message, subject) {
  var ADMIN_EMAIL = "mhmdfdln14@gmail.com";
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject || "Approval System Notification",
      body: message
    });
    Logger.log("Admin notification sent to: " + ADMIN_EMAIL);
  } catch (error) {
    Logger.log("Failed to send admin notification: " + error.toString());
  }
}

function sendNextApprovalAfterFirstLayer() {
  Logger.log("Checking for Second Layer approvals after First Layer...");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
  var processedCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var firstLayerStatus = row[7]; // Column H
    var secondLayerStatus = row[8]; // Column I
    var currentStatus = row[13]; // Column N

    Logger.log("Row " + (i+2) + ": First=" + firstLayerStatus + ", Second=" + secondLayerStatus + ", Status=" + currentStatus);

    if (firstLayerStatus === "APPROVED" && (!secondLayerStatus || secondLayerStatus === "" || secondLayerStatus === "PENDING") && currentStatus === "PROCESSING") {
      Logger.log(" Found eligible for Second Layer approval: " + row[0]);

      var secondLayerEmail = row[11]; // Column L
      var name = row[0];
      var email = row[1];
      var description = row[2];
      var documentType = row[3];
      var attachment = row[4];

      if (secondLayerEmail) {
        var validationResult = validateGoogleDriveAttachmentWithType(attachment, documentType);
        var approvalLink = generateMultiLayerApprovalLink(name, email, description, documentType, attachment, "SECOND_LAYER");
        var emailSent = sendMultiLayerEmail(secondLayerEmail, description, documentType, attachment, approvalLink, "SECOND_LAYER", validationResult);

        if (emailSent) {
          sheet.getRange(i + 2, 7).setNote("SECOND_LAYER_APPROVAL_SENT: " + new Date());
          Logger.log(" Second Layer approval email sent to: " + secondLayerEmail);
          processedCount++;
          Utilities.sleep(1000);
        }
      }
    }
  }

  if (processedCount > 0) {
    Logger.log("3 daysSuccessfully sent " + processedCount + " Second Layer approval emails!");
  } else {
    Logger.log("No pending Second Layer approvals found");
  }
}

function sendNextApprovalAfterSecondLayer() {
  Logger.log("Checking for Third Layer approvals after Second Layer...");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
  var processedCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var secondLayerStatus = row[8]; // Column I
    var thirdLayerStatus = row[9]; // Column J
    var currentStatus = row[13]; // Column N

    Logger.log("Row " + (i+2) + ": Second=" + secondLayerStatus + ", Third=" + thirdLayerStatus + ", Status=" + currentStatus);

    if (secondLayerStatus === "APPROVED" && (!thirdLayerStatus || thirdLayerStatus === "" || thirdLayerStatus === "PENDING") && currentStatus === "PROCESSING") {
      Logger.log(" Found eligible for Third Layer approval: " + row[0]);

      var thirdLayerEmail = row[12]; // Column M
      var name = row[0];
      var email = row[1];
      var description = row[2];
      var documentType = row[3];
      var attachment = row[4];

      if (thirdLayerEmail) {
        var validationResult = validateGoogleDriveAttachmentWithType(attachment, documentType);
        var approvalLink = generateMultiLayerApprovalLink(name, email, description, documentType, attachment, "THIRD_LAYER");
        var emailSent = sendMultiLayerEmail(thirdLayerEmail, description, documentType, attachment, approvalLink, "THIRD_LAYER", validationResult);

        if (emailSent) {
          sheet.getRange(i + 2, 7).setNote("THIRD_LAYER_APPROVAL_SENT: " + new Date());
          Logger.log(" Third Layer approval email sent to: " + thirdLayerEmail);
          processedCount++;
          Utilities.sleep(1000);
        }
      }
    }
  }

  if (processedCount > 0) {
    Logger.log("3 daysSuccessfully sent " + processedCount + " Third Layer approval emails!");
  } else {
    Logger.log("No pending Third Layer approvals found");
  }
}

// ============================================
// WEB APP HANDLERS
// ============================================

function doGet(e) {
  try {
    Logger.log("Web App accessed");
    Logger.log("e.parameter: " + JSON.stringify(e.parameter));
    
    var params = e.parameter || {};
    var action = params.action;
    
    if (action === "approve") {
      return handleMultiLayerApproval(params);
    } else if (action === "reject") {
      return handleMultiLayerRejection(params);
    } else if (action === "rejection_success") {
      return showRejectionSuccessPage(params);
    }
    
    return createErrorPage("Invalid request - missing action parameter");
    
  } catch (error) {
    Logger.log("Error in doGet: " + error.toString());
    return createErrorPage("System error: " + error.message);
  }
}

function handleMultiLayerApproval(params) {
  try {
    Logger.log("Approval request received");
    
    var name = params.name ? decodeURIComponent(params.name) : "";
    var email = params.email ? decodeURIComponent(params.email) : "";
    var project = params.project ? decodeURIComponent(params.project) : "";
    var documentType = params.docType ? decodeURIComponent(params.docType) : "";
    var attachment = params.attachment ? decodeURIComponent(params.attachment) : "";
    var layer = params.layer ? decodeURIComponent(params.layer) : "";
    var code = params.code || "";
    var timestamp = parseInt(params.timestamp) || 0;
    
    var now = new Date().getTime();
    var sevenDaysAgo = now - (7 * 24 * 60 * 90 * 1000);
    if (timestamp < sevenDaysAgo) {
      return createErrorPage("Approval link has expired (3 days). Please request a new one.");
    }
    
    if (!name || !email || !project || !layer) {
      return createErrorPage("Missing required approval data");
    }
    
    var updated = updateMultiLayerApprovalStatus(name, email, project, layer, code);
    
    if (updated) {
      Logger.log(" Approval updated successfully");
      
      try {
        if (layer === "FIRST_LAYER") {
          Utilities.sleep(3000);
          sendNextApprovalAfterFirstLayer();
        } else if (layer === "SECOND_LAYER") {
          Utilities.sleep(3000);
          sendNextApprovalAfterSecondLayer();
        }
      } catch (nextError) {
        Logger.log("Warning: Next approval trigger failed: " + nextError.toString());
      }
      
      return createSuccessPage(name, email, project, documentType, attachment, layer, code);
    } else {
      return createErrorPage("Approval failed - data not found or already approved");
    }
    
  } catch (error) {
    Logger.log("Error in handleMultiLayerApproval: " + error.toString());
    return createErrorPage("System error during approval: " + error.message);
  }
}

function updateMultiLayerApprovalStatus(name, email, project, layer, code) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return false;
    }
    
    var data = sheet.getRange("A2:N" + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowName = row[0];
      var rowEmail = row[1];
      var rowProject = row[2];
      var rowStatus = row[13]; // Column N - Status
      
      // Skip empty rows
      if (!rowName && !rowEmail && !rowProject) continue;
      
      // FIXED: SUPER STRICT MATCHING - harus exact match semua field yang ada
      var nameMatch = rowName && name && 
                     rowName.toString().trim().toLowerCase() === name.toString().trim().toLowerCase();
      
      var emailMatch = rowEmail && email && 
                      rowEmail.toString().trim().toLowerCase() === email.toString().trim().toLowerCase();
      
      var projectMatch = rowProject && project && 
                        rowProject.toString().trim().toLowerCase() === project.toString().trim().toLowerCase();
      
      // FIXED: Hanya proses row yang statusnya PROCESSING/ACTIVE dan layer yang sesuai masih PENDING
      var isEligibleForUpdate = (rowStatus === "PROCESSING" || rowStatus === "ACTIVE");
      
      // Check layer status - hanya proses jika layer tersebut masih PENDING
      var layerColumn = getLayerColumnIndex(layer);
      var currentLayerStatus = layerColumn !== -1 ? sheet.getRange(i + 2, layerColumn).getValue() : "";
      var isLayerPending = (!currentLayerStatus || currentLayerStatus === "" || currentLayerStatus === "PENDING");
      
      // FIXED: HARUS MATCH SEMUA 3 FIELD + ELIGIBLE + LAYER PENDING
      if (nameMatch && emailMatch && projectMatch && isEligibleForUpdate && isLayerPending) {
        Logger.log("✅ EXACT MATCH FOUND at row " + (i + 2));
        Logger.log("   Name: '" + rowName + "' = '" + name + "'");
        Logger.log("   Email: '" + rowEmail + "' = '" + email + "'");
        Logger.log("   Project: '" + rowProject + "' = '" + project + "'");
        Logger.log("   Status: " + rowStatus + ", Layer Status: " + currentLayerStatus);
        
        var columnIndex = getLayerColumnIndex(layer);
        if (columnIndex !== -1) {
          var currentStatus = sheet.getRange(i + 2, columnIndex).getValue();
          if (currentStatus === "APPROVED" || currentStatus === "REJECTED") {
            Logger.log("⚠️ Layer already processed - skipping");
            return false;
          }
          
          // UPDATE THE CORRECT ROW
          sheet.getRange(i + 2, columnIndex).setValue("APPROVED");
          sheet.getRange(i + 2, columnIndex).setBackground("#90EE90");
          sheet.getRange(i + 2, columnIndex).setNote("Approved by " + getLayerDisplayName(layer) + " - " + getGMT7Time() + " - Code: " + code);
          
          // Check if all layers are approved
          var firstLayerStatus = sheet.getRange(i + 2, 8).getValue(); // Column H
          var secondLayerStatus = sheet.getRange(i + 2, 9).getValue(); // Column I
          var thirdLayerStatus = sheet.getRange(i + 2, 10).getValue(); // Column J

          if (firstLayerStatus === "APPROVED" && secondLayerStatus === "APPROVED" && thirdLayerStatus === "APPROVED") {
            sheet.getRange(i + 2, 14).setValue("COMPLETED"); // Column N
            sheet.getRange(i + 2, 14).setBackground("#90EE90");
            Logger.log("✅ All layers approved - marked as COMPLETED");
          }
          
          return true;
        }
      }
    }
    
    Logger.log("❌ No exact matching data found");
    Logger.log("   Searching for: Name='" + name + "', Email='" + email + "', Project='" + project + "'");
    
    // FIXED: FALLBACK - coba dengan matching yang sedikit lebih longgar tapi masih strict
    return fallbackStrictMatch(sheet, data, name, email, project, layer, code, "APPROVED");
    
  } catch (error) {
    Logger.log("❌ Error updating status: " + error.toString());
    return false;
  }
}

// NEW FUNCTION: Fallback matching yang masih strict
function fallbackStrictMatch(sheet, data, name, email, project, layer, code, action) {
  Logger.log("🔄 Trying fallback strict matching...");
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowName = row[0];
    var rowEmail = row[1];
    var rowProject = row[2];
    var rowStatus = row[13];
    
    if (!rowName && !rowEmail && !rowProject) continue;
    
    // FALLBACK: Project harus exact match + salah satu dari name/email match
    var nameMatch = rowName && name && 
                   rowName.toString().trim().toLowerCase() === name.toString().trim().toLowerCase();
    
    var emailMatch = rowEmail && email && 
                    rowEmail.toString().trim().toLowerCase() === email.toString().trim().toLowerCase();
    
    var projectMatch = rowProject && project && 
                      rowProject.toString().trim().toLowerCase() === project.toString().trim().toLowerCase();
    
    var isEligibleForUpdate = (rowStatus === "PROCESSING" || rowStatus === "ACTIVE");
    
    // Check layer status
    var layerColumn = getLayerColumnIndex(layer);
    var currentLayerStatus = layerColumn !== -1 ? sheet.getRange(i + 2, layerColumn).getValue() : "";
    var isLayerPending = (!currentLayerStatus || currentLayerStatus === "" || currentLayerStatus === "PENDING");
    
    // FALLBACK: Project exact match + (name match ATAU email match) + eligible + layer pending
    if (projectMatch && isEligibleForUpdate && isLayerPending && (nameMatch || emailMatch)) {
      Logger.log("✅ FALLBACK MATCH FOUND at row " + (i + 2));
      Logger.log("   Project exact match + " + (nameMatch ? "Name match" : "Email match"));
      
      var columnIndex = getLayerColumnIndex(layer);
      if (columnIndex !== -1) {
        var currentStatus = sheet.getRange(i + 2, columnIndex).getValue();
        if (currentStatus === "APPROVED" || currentStatus === "REJECTED") {
          Logger.log("⚠️ Layer already processed - skipping");
          return false;
        }
        
        if (action === "APPROVED") {
          sheet.getRange(i + 2, columnIndex).setValue("APPROVED");
          sheet.getRange(i + 2, columnIndex).setBackground("#90EE90");
          sheet.getRange(i + 2, columnIndex).setNote("Approved by " + getLayerDisplayName(layer) + " - " + getGMT7Time() + " - Code: " + code);
        } else {
          sheet.getRange(i + 2, columnIndex).setValue("REJECTED");
          sheet.getRange(i + 2, columnIndex).setBackground("#FF6B6B");
          sheet.getRange(i + 2, columnIndex).setNote("Rejected by " + getLayerDisplayName(layer) + " - " + getGMT7Time() + " - Code: " + code);
        }
        
        return true;
      }
    }
  }
  
  Logger.log("❌ No fallback match found either");
  return false;
}

function getLayerColumnIndex(layer) {
  var layerColumns = {
    "FIRST_LAYER": 8,  // Column H
    "SECOND_LAYER": 9, // Column I
    "THIRD_LAYER": 10  // Column J
  };
  return layerColumns[layer] || -1;
}

// ============================================
// REJECTION HANDLING
// ============================================

function doPost(e) {
  try {
    Logger.log("POST request received for rejection");
    
    var params = {};
    
    // Handle form data properly
    if (e.postData && e.postData.contents) {
      var contents = e.postData.contents;
      Logger.log("Raw POST data: " + contents);
      
      // Parse form data
      var formData = contents.split('&');
      for (var i = 0; i < formData.length; i++) {
        var pair = formData[i].split('=');
        if (pair.length === 2) {
          params[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1]);
        }
      }
    }
    
    Logger.log("Processed POST params: " + JSON.stringify(params));
    
    var action = params.action;
    Logger.log("POST action: " + action);

    if (action === "submit_rejection") {
      return handleRejectionSubmission(params);
    }

    return createErrorPage("Invalid POST request");

  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    return createErrorPage("System error: " + error.message);
  }
}

// ============================================
// FIXED REJECTION FUNCTIONS
// ============================================

function handleMultiLayerRejection(params) {
  try {
    Logger.log("Rejection request received");
    Logger.log("All rejection params: " + JSON.stringify(params));

    // FIXED: Safe parameter extraction with better decoding
    var name = params.name ? decodeURIComponent(params.name) : (params.name || "");
    var email = params.email ? decodeURIComponent(params.email) : (params.email || "");
    var project = params.project ? decodeURIComponent(params.project) : (params.project || "");
    var documentType = params.docType ? decodeURIComponent(params.docType) : (params.docType || "");
    var attachment = params.attachment ? decodeURIComponent(params.attachment) : (params.attachment || "");
    var layer = params.layer ? decodeURIComponent(params.layer) : (params.layer || "");
    var code = params.code || "";
    var timestamp = parseInt(params.timestamp) || 0;

    // FIXED: Debug logging untuk semua parameter
    Logger.log("Rejection Parameter Debug:");
    Logger.log("Name: '" + name + "'");
    Logger.log("Email: '" + email + "'");
    Logger.log("Project: '" + project + "'");
    Logger.log("Layer: '" + layer + "'");
    Logger.log("DocumentType: '" + documentType + "'");
    Logger.log("Attachment: '" + attachment + "'");
    Logger.log("Code: '" + code + "'");

    // FIXED: Check link expiration
    var now = new Date().getTime();
    var sevenDaysAgo = now - (7 * 24 * 60 * 60 * 1000);
    if (timestamp < sevenDaysAgo) {
      return createErrorPage("Rejection link has expired (7 days). Please request a new one.");
    }

    // FIXED: More flexible validation - cari data dari spreadsheet jika parameter kosong
    if (!name || !project || !layer) {
      Logger.log("Missing parameters, trying to find data from spreadsheet...");
      
      // Coba cari data dari spreadsheet berdasarkan parameter yang ada
      var foundData = findRejectionDataFromSpreadsheet(email, project, layer);
      if (foundData) {
        name = foundData.name;
        project = foundData.project;
        layer = foundData.layer;
        email = foundData.email;
        Logger.log("Data found in spreadsheet: " + name + " - " + project);
      } else {
        var missingFields = [];
        if (!name) missingFields.push("name");
        if (!project) missingFields.push("project"); 
        if (!layer) missingFields.push("layer");
        
        Logger.log("Missing fields after search: " + missingFields.join(", "));
        return createErrorPage("Cannot process rejection: Missing " + missingFields.join(", ") + ". Please use the original approval link from email.");
      }
    }

    // FIXED: Create rejection form dengan data yang sudah diverifikasi
    return createRejectionForm(name, email, project, documentType, attachment, layer, code);

  } catch (error) {
    Logger.log("Error in handleMultiLayerRejection: " + error.toString());
    return createErrorPage("System error during rejection: " + error.message);
  }
}

// NEW FUNCTION: Cari data dari spreadsheet jika parameter kosong
function findRejectionDataFromSpreadsheet(email, project, layer) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowEmail = row[1]; // Column B
      var rowProject = row[2]; // Column C
      var rowStatus = row[13]; // Column N - Status
      
      // FIXED: SUPER STRICT MATCHING
      var emailMatch = rowEmail && email && 
                      rowEmail.toString().trim().toLowerCase() === email.toString().trim().toLowerCase();
      var projectMatch = rowProject && project && 
                        rowProject.toString().trim().toLowerCase() === project.toString().trim().toLowerCase();
      var isEligible = (rowStatus === "PROCESSING" || rowStatus === "ACTIVE");
      
      if (emailMatch && projectMatch && isEligible) {
        Logger.log("✅ Found exact matching data in row " + (i + 2));
        return {
          name: row[0], // Column A
          email: row[1], // Column B
          project: row[2], // Column C
          layer: layer,
          documentType: row[3], // Column D
          attachment: row[4] // Column E
        };
      }
    }
    
    Logger.log("❌ No exact matching data found in spreadsheet");
    return null;
    
  } catch (error) {
    Logger.log("❌ Error finding rejection data: " + error.toString());
    return null;
  }
}

// NEW FUNCTION: Get layer status from row
function getLayerStatusFromRow(row, layer) {
  if (layer === "FIRST_LAYER") return row[7]; // Column H
  if (layer === "SECOND_LAYER") return row[8]; // Column I
  if (layer === "THIRD_LAYER") return row[9]; // Column J
  return "";
}

// FIXED: Update juga function generateMultiLayerApprovalLink untuk rejection
function generateMultiLayerApprovalLink(name, email, description, documentType, attachment, layer, action) {
  var webAppUrl = "https://script.google.com/macros/s/AKfycbw-7cNeUsM82jvqztSJ4GpSI4Nzdg2Fq-yS-2kAGDkUyKF-uvX3MDVsvQXuqNKcxBd-/exec";

  var timestamp = new Date().getTime().toString(36);
  var nameCode = name ? name.substring(0, 3).toLowerCase() : "usr";
  var projectCode = description ? description.replace(/[^a-zA-Z0-9]/g, "").substring(0, 3).toLowerCase() : "prj";
  var uniqueCode = nameCode + projectCode + timestamp;

  // FIXED: Pastikan semua parameter ada dan properly encoded
  var params = {
    action: action || "approve",
    name: encodeURIComponent(name || "Unknown"),
    email: encodeURIComponent(email || "unknown@atreusg.com"),
    project: encodeURIComponent(description || "Unknown Project"),
    docType: encodeURIComponent(documentType || ""),
    attachment: encodeURIComponent(attachment || ""),
    layer: layer,
    code: uniqueCode,
    timestamp: new Date().getTime()
  };

  // FIXED: Debug logging untuk link generation
  Logger.log("Generated " + (action || "approve") + " link for: " + name + " - " + description);
  Logger.log("Link params: " + JSON.stringify(params));

  var queryString = Object.keys(params)
    .map(key => key + '=' + params[key])
    .join('&');

  return webAppUrl + "?" + queryString;
}

// FIXED: Update juga di email function untuk rejection link
function sendMultiLayerEmail(recipientEmail, description, documentType, attachment, approvalLink, layer, validationResult) {
  try {
    if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
      Logger.log("Invalid email: " + recipientEmail);
      return false;
    }
    
    var layerDisplayNames = {
      "FIRST_LAYER": "First Layer",
      "SECOND_LAYER": "Second Layer",
      "THIRD_LAYER": "Third Layer"
    };
    
    var layerDisplay = layerDisplayNames[layer] || layer;
    var companyName = "Atreus Global";
    
    var subject = "Approval Request - " + description + " [" + layerDisplay + " Layer]";
    
    var docTypeBadge = "";
    if (documentType && documentType !== "") {
      var docTypeColors = {
        "ICC": "#3B82F6",
        "Quotation": "#10B981",
        "Proposal": "#F59E0B"
      };
      var badgeColor = docTypeColors[documentType] || "#6B7280";
      docTypeBadge = '<div style="display: inline-block; background: ' + badgeColor + '; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-bottom: 10px;">' + documentType + '</div>';
    }
    
    var attachmentSection = "";
    if (attachment && attachment !== "") {
      if (validationResult.valid) {
        var driveTypeInfo = validationResult.isSharedDrive ? 
          '<p><strong>Location:</strong> <span style="color: #3B82F6;">📁 Shared Drive</span></p>' :
          '<p><strong>Location:</strong> My Drive</p>';
        
        attachmentSection = '<div class="attachment-box" style="background: linear-gradient(135deg, #f0f8ff 0%, #e6f2ff 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #90EE90;"><h3 style="color: #2E8B57;">Attachment (Validated)</h3>' + docTypeBadge + '<p><strong>File:</strong> ' + validationResult.name + '</p><p><strong>Type:</strong> ' + validationResult.type + '</p><p><strong>Size:</strong> ' + validationResult.sizeFormatted + '</p>' + driveTypeInfo + '<p><strong>Status:</strong> <span style="color: #2E8B57;">' + validationResult.message + '</span></p><p><a href="' + attachment + '" target="_blank" style="color: #326BC6;">View Document</a></p></div>';
      } else {
        attachmentSection = '<div class="attachment-box" style="background: linear-gradient(135deg, #fff0f0 0%, #ffe6e6 100%); padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #FF6B6B;"><h3 style="color: #DC143C;">Attachment (Validation Failed)</h3>' + docTypeBadge + '<p><strong>Status:</strong> <span style="color: #DC143C;">' + validationResult.message + '</span></p><p><a href="' + attachment + '" target="_blank" style="color: #326BC6;">Check Link</a></p><p style="font-size: 12px; color: #666;"><em>Please ensure the Google Drive link is correct and accessible</em></p></div>';
      }
    }
    
    var progressBar = getProgressBar(layer);
    
    // FIXED: Generate rejection link dengan parameter yang sama seperti approval
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
    
    // Cari data yang sesuai untuk rejection link
    var rejectionName = "Unknown";
    var rejectionEmail = "unknown@atreusg.com";
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row[2] === description) { // Match berdasarkan project description
        rejectionName = row[0] || "Unknown";
        rejectionEmail = row[1] || "unknown@atreusg.com";
        break;
      }
    }
    
    var rejectLink = generateMultiLayerApprovalLink(rejectionName, rejectionEmail, description, documentType, attachment, layer, "reject");
    
    Logger.log("Rejection link generated for: " + rejectionName + " - " + description);
    
    var htmlBody = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;line-height:1.6;color:#333;background:#f6f9fc;margin:0;padding:0}.container{max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1)}.header{background:linear-gradient(135deg,#326BC6 0%,#183460 100%);padding:30px;text-align:center;color:white}.progress-track{background:#f8f9fa;padding:30px 20px;margin:0}.progress-steps{display:flex;justify-content:center;align-items:center;gap:40px;position:relative}.progress-step{text-align:center;position:relative}.step-number{width:40px;height:40px;border-radius:50%;margin:0 auto 8px;font-weight:600;display:flex;align-items:center;justify-content:center;font-size:16px}.step-active{background:#326BC6;color:white}.step-completed{background:#90EE90;color:#333}.step-pending{background:#e0e0e0;color:#666}.step-label{font-size:12px;font-weight:500}.content{padding:30px}.button{display:inline-block;padding:15px 30px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white !important;text-decoration:none;border-radius:8px;margin:20px 10px;font-weight:600;font-size:16px;border:none;cursor:pointer;transition:all 0.3s ease}.button:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(50,107,198,0.3)}.button-reject{background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white !important;text-decoration:none;border-radius:8px;margin:20px 10px;font-weight:600;font-size:16px;border:none;cursor:pointer;transition:all 0.3s ease}.button-reject:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(220,38,38,0.3)}.button-container{text-align:center;margin:20px 0}.info-box{background:#f8f9fa;padding:15px;border-radius:5px;margin:15px 0;border-left:4px solid #326BC6}.footer{margin-top:30px;padding:20px;background:#f8f9fa;text-align:center;font-size:12px;color:#666}</style></head><body><div class="container"><div class="header"><h1>Multi-Layer Approval Required</h1><p>Current Stage: ' + layerDisplay + ' Approval</p></div><div class="progress-track">' + progressBar + '</div><div class="content"><p>Hello,</p><p>This project requires your approval at the <strong>' + layerDisplay + '</strong> level:</p><div class="info-box"><h3>' + description + '</h3><p><strong>Current Stage:</strong> ' + layerDisplay + ' Approval</p><p><strong>Date:</strong> ' + getGMT7Time() + '</p></div>' + attachmentSection + '<p>Please choose your decision:</p><div class="button-container"><a href="' + approvalLink + '" class="button" style="color: white !important;">APPROVE AS ' + layerDisplay + '</a><a href="' + rejectLink + '" class="button-reject" style="color: white !important;">REJECT WITH NOTE</a></div><p style="text-align: center; font-size: 12px; color: #666;"><em>Approval Code: ' + getApprovalCodeFromLink(approvalLink) + ' | Link expires in 7 days</em></p></div><div class="footer"><p>This is an automated email. Please do not reply.</p><p>© ' + new Date().getFullYear() + ' ' + companyName + '. All rights reserved.</p></div></div></body></html>';
    
    var plainBody = "MULTI-LAYER APPROVAL REQUEST\n\nProject: " + description + "\nDocument Type: " + documentType + "\nCurrent Stage: " + layerDisplay + " Approval\n\n";
    
    if (attachment && attachment !== "") {
      plainBody += "Attachment: " + attachment + "\nAttachment Status: " + validationResult.message + "\n\n";
    }
    
    plainBody += "Approve: " + approvalLink + "\nReject: " + rejectLink + "\n\nThank you.\n\nThis is an automated email from " + companyName + ".";
    
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody
    });
    
    Logger.log("Multi-layer email sent to: " + recipientEmail + " for layer: " + layerDisplay);
    return true;
    
  } catch (error) {
    Logger.log("Error sending multi-layer email: " + error.toString());
    return false;
  }
}

function createRejectionForm(name, email, project, documentType, attachment, layer, code) {
  var layerDisplayNames = {
    "FIRST_LAYER": "First Layer",
    "SECOND_LAYER": "Second Layer",
    "THIRD_LAYER": "Third Layer"
  };

  var layerDisplay = layerDisplayNames[layer] || layer;

  var docTypeBadge = "";
  if (documentType && documentType !== "") {
    var docTypeColors = {
      "ICC": "#3B82F6",
      "Quotation": "#10B981",
      "Proposal": "#F59E0B"
    };
    var badgeColor = docTypeColors[documentType] || "#6B7280";
    docTypeBadge = '<div style="display: inline-block; background: ' + badgeColor + '; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; margin: 10px 0;">' + documentType + '</div>';
  }

  // FIXED: SIMPLE HTML dengan JavaScript yang LENGKAP dan WORK
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
    .loading {
      display: none;
      color: #dc2626;
      font-weight: 600;
      margin-top: 15px;
    }
    .success-message {
      color: #059669;
      font-weight: 600;
      margin-top: 15px;
    }
    .error-message {
      color: #dc2626;
      font-weight: 600;
      margin-top: 15px;
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

    <form id="rejectionForm">
      <input type="hidden" name="name" value="${name || ''}">
      <input type="hidden" name="email" value="${email || ''}">
      <input type="hidden" name="project" value="${project || ''}">
      <input type="hidden" name="docType" value="${documentType || ''}">
      <input type="hidden" name="attachment" value="${attachment || ''}">
      <input type="hidden" name="layer" value="${layer || ''}">
      <input type="hidden" name="code" value="${code || ''}">
      
      <div class="rejection-form">
        <div class="form-group">
          <label class="form-label" for="rejectionNote">Rejection Reason (Required):</label>
          <textarea id="rejectionNote" name="rejectionNote" class="form-textarea" 
                    placeholder="Please explain why this request is being rejected. This note will be visible to the requester and previous approvers." 
                    required></textarea>
        </div>
      </div>
      
      <div style="text-align: center;">
        <button type="button" id="submitBtn" class="button" onclick="submitRejectionForm()">Submit Rejection</button>
      </div>
      
      <div id="loading" class="loading">Processing your rejection... Please wait.</div>
      <div id="successMessage" class="success-message" style="display: none;">Rejection submitted successfully! Redirecting...</div>
      <div id="errorMessage" class="error-message" style="display: none;"></div>
    </form>
    
    <div class="company-brand">
      <strong>Atreus Global</strong> • Multi-Layer Approval System
    </div>
  </div>

  <script>
    // FIXED: COMPLETE JavaScript function - TIDAK ADA YANG TERPOTONG
    function submitRejectionForm() {
      console.log('Submit button clicked');
      
      var submitBtn = document.getElementById('submitBtn');
      var loading = document.getElementById('loading');
      var successMessage = document.getElementById('successMessage');
      var errorMessage = document.getElementById('errorMessage');
      var rejectionNote = document.getElementById('rejectionNote').value;
      
      // Reset messages
      errorMessage.style.display = 'none';
      successMessage.style.display = 'none';
      
      // Validation
      if (!rejectionNote.trim()) {
        errorMessage.textContent = 'Please provide a rejection reason.';
        errorMessage.style.display = 'block';
        return;
      }
      
      // Show loading, disable button
      submitBtn.disabled = true;
      loading.style.display = 'block';
      
      // Prepare form data
      var form = document.getElementById('rejectionForm');
      var formData = new FormData(form);
      var params = {};
      
      for (var pair of formData.entries()) {
        params[pair[0]] = pair[1];
      }
      
      params.rejectionNote = rejectionNote;
      params.action = 'submit_rejection';
      
      console.log('Submitting rejection with params:', params);
      
      // FIXED: Use google.script.run dengan error handling yang proper
      google.script.run
        .withSuccessHandler(function(result) {
          console.log('Rejection success:', result);
          loading.style.display = 'none';
          successMessage.style.display = 'block';
          
          // Redirect ke halaman success
          setTimeout(function() {
            // Redirect ke success page
            var successUrl = 'https://script.google.com/macros/s/AKfycbyfz6LxLBXCIoOLcIty4Kcwq61Jj9zYUkzqHeEEWVX_3pOln796hX9gueXFw0l4PPbd/exec' +
              '?action=rejection_success' +
              '&project=' + encodeURIComponent(params.project || '') +
              '&layer=' + encodeURIComponent(params.layer || '') +
              '&name=' + encodeURIComponent(params.name || '');
            window.location.href = successUrl;
          }, 1500);
        })
        .withFailureHandler(function(error) {
          console.error('Rejection failed:', error);
          loading.style.display = 'none';
          submitBtn.disabled = false;
          errorMessage.textContent = 'Rejection failed: ' + error.message;
          errorMessage.style.display = 'block';
        })
        .handleRejectionSubmission(params);
    }
    
    // FIXED: Add event listener untuk Enter key
    document.getElementById('rejectionNote').addEventListener('keydown', function(e) {
      if (e.ctrlKey && e.key === 'Enter') {
        submitRejectionForm();
      }
    });
    
    console.log('Rejection form loaded successfully');
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleRejectionSubmission(params) {
    try {
        Logger.log("Rejection submission received - FIXED VERSION");
        
        // Extract parameters
        var name = params.name || "";
        var email = params.email || "";
        var project = params.project || "";
        var documentType = params.docType || "";
        var attachment = params.attachment || "";
        var layer = params.layer || "";
        var code = params.code || "";
        var rejectionNote = params.rejectionNote || "";

        Logger.log("Rejection data - Name: " + name + ", Project: " + project + ", Layer: " + layer);

        // Validation
        if (!rejectionNote || rejectionNote.trim() === "") {
            return createErrorPage("Rejection note is required.");
        }

        // Update spreadsheet
        var updated = updateMultiLayerRejectionStatus(name, email, project, layer, code, rejectionNote);

        if (updated) {
            Logger.log("Rejection updated successfully");

            // Send notifications
            try {
                sendRejectionNotification(name, email, project, documentType, attachment, layer, rejectionNote);
                Logger.log("Notifications sent");
            } catch (notifyError) {
                Logger.log("Notification failed: " + notifyError.toString());
            }

            // RETURN SUCCESS PAGE - PASTI WORK
            return createRejectionSuccessPage(name, email, project, documentType, attachment, layer, rejectionNote);
            
        } else {
            Logger.log("Rejection failed - data not found");
            return createErrorPage("Rejection failed - data not found in spreadsheet");
        }

    } catch (error) {
        Logger.log("Error in handleRejectionSubmission: " + error.toString());
        return createErrorPage("System error during rejection: " + error.message);
    }
}

function updateMultiLayerRejectionStatus(name, email, project, layer, code, rejectionNote) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return false;
    }

    var data = sheet.getRange("A2:N" + lastRow).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowName = row[0];
      var rowEmail = row[1];
      var rowProject = row[2];
      var rowStatus = row[13];

      if (!rowName && !rowEmail && !rowProject) continue;

      // FIXED: SUPER STRICT MATCHING - sama seperti approval
      var nameMatch = rowName && name && 
                     rowName.toString().trim().toLowerCase() === name.toString().trim().toLowerCase();
      
      var emailMatch = rowEmail && email && 
                      rowEmail.toString().trim().toLowerCase() === email.toString().trim().toLowerCase();
      
      var projectMatch = rowProject && project && 
                        rowProject.toString().trim().toLowerCase() === project.toString().trim().toLowerCase();

      // FIXED: Hanya proses row yang statusnya PROCESSING/ACTIVE dan layer yang sesuai masih PENDING
      var isEligibleForUpdate = (rowStatus === "PROCESSING" || rowStatus === "ACTIVE");
      
      // Check layer status
      var layerColumn = getLayerColumnIndex(layer);
      var currentLayerStatus = layerColumn !== -1 ? sheet.getRange(i + 2, layerColumn).getValue() : "";
      var isLayerPending = (!currentLayerStatus || currentLayerStatus === "" || currentLayerStatus === "PENDING");

      // FIXED: HARUS MATCH SEMUA 3 FIELD + ELIGIBLE + LAYER PENDING
      if (nameMatch && emailMatch && projectMatch && isEligibleForUpdate && isLayerPending) {
        Logger.log("✅ EXACT MATCH FOUND at row " + (i + 2) + " for rejection");
        Logger.log("   Name: '" + rowName + "' = '" + name + "'");
        Logger.log("   Email: '" + rowEmail + "' = '" + email + "'");
        Logger.log("   Project: '" + rowProject + "' = '" + project + "'");

        var columnIndex = getLayerColumnIndex(layer);
        if (columnIndex !== -1) {
          var currentStatus = sheet.getRange(i + 2, columnIndex).getValue();
          if (currentStatus === "REJECTED" || currentStatus === "APPROVED") {
            Logger.log("⚠️ Layer already processed - skipping");
            return false;
          }

          // Set current layer to REJECTED
          sheet.getRange(i + 2, columnIndex).setValue("REJECTED");
          sheet.getRange(i + 2, columnIndex).setBackground("#FF6B6B");
          sheet.getRange(i + 2, columnIndex).setNote("Rejected by " + getLayerDisplayName(layer) + " - " + getGMT7Time() + " - Code: " + code + "\nReason: " + rejectionNote);

          // RESET SEMUA LAYER SETELAHNYA KE PENDING
          if (layer === "FIRST_LAYER") {
            sheet.getRange(i + 2, 9).setValue("PENDING");  // Second Layer
            sheet.getRange(i + 2, 9).setBackground(null);
            sheet.getRange(i + 2, 9).setNote("");
            sheet.getRange(i + 2, 10).setValue("PENDING"); // Third Layer
            sheet.getRange(i + 2, 10).setBackground(null);
            sheet.getRange(i + 2, 10).setNote("");
          } else if (layer === "SECOND_LAYER") {
            sheet.getRange(i + 2, 10).setValue("PENDING"); // Third Layer
            sheet.getRange(i + 2, 10).setBackground(null);
            sheet.getRange(i + 2, 10).setNote("");
          }

          // SET OVERALL STATUS KE REJECTED
          sheet.getRange(i + 2, 14).setValue("REJECTED");
          sheet.getRange(i + 2, 14).setBackground("#FF6B6B");
          
          // UPDATE LOG NOTES
          var currentNote = sheet.getRange(i + 2, 7).getNote();
          if (currentNote && currentNote.includes("APPROVAL_LINK")) {
            sheet.getRange(i + 2, 7).setNote("REJECTED: " + getGMT7Time() + " - Reason: " + rejectionNote);
          }

          Logger.log("✅ Rejection recorded for layer: " + layer + " at row " + (i + 2));
          return true;
        }
      }
    }

    Logger.log("❌ No exact matching data found for rejection");
    
    // FIXED: FALLBACK - sama seperti approval
    return fallbackStrictMatch(sheet, data, name, email, project, layer, code, "REJECTED");

  } catch (error) {
    Logger.log("❌ Error updating rejection status: " + error.toString());
    return false;
  }
}

// FIXED: IMPROVED REJECTION NOTIFICATION - SESUAI FLOW DIAGRAM
function sendRejectionNotification(name, email, project, documentType, attachment, layer, rejectionNote) {
  try {
    var layerDisplayNames = {
      "FIRST_LAYER": "First Layer",
      "SECOND_LAYER": "Second Layer", 
      "THIRD_LAYER": "Third Layer"
    };

    var layerDisplay = layerDisplayNames[layer] || layer;
    var companyName = "Atreus Global";
    var currentTime = getGMT7Time();

    // 1. NOTIFY REQUESTER (SELALU)
    var requesterSubject = "Approval Request Rejected - " + project;
    var requesterBody = "Dear " + (name || "Requester") + ",\n\n" +
      "Your approval request has been REJECTED at the " + layerDisplay + " level.\n\n" +
      "Project: " + project + "\n" +
      "Rejected at: " + layerDisplay + " Approval\n" +
      "Rejection Date: " + currentTime + "\n\n" +
      "REJECTION REASON:\n" + rejectionNote + "\n\n" +
      "NEXT STEPS:\n" +
      "Please review the feedback above and make necessary changes to your submission.\n" +
      "Once revised, you may resubmit for approval.\n\n" +
      "Best regards,\n" +
      companyName + " Approval System";

    MailApp.sendEmail({
      to: email,
      subject: requesterSubject,
      body: requesterBody
    });

    Logger.log("Rejection notification sent to requester: " + email);

    // 2. NOTIFY PREVIOUS APPROVERS (jika ada)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
    var previousApprovers = [];

    // Cari row yang sesuai untuk dapat email previous approvers
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowName = row[0];
      var rowEmail = row[1];
      var rowProject = row[2];

      if (!rowName || !rowProject) continue;

      var nameMatch = rowName.toString().trim().toLowerCase() === (name || "").toString().trim().toLowerCase();
      var projectMatch = rowProject.toString().trim().toLowerCase().includes((project || "").toString().trim().toLowerCase());

      if (nameMatch && projectMatch) {
        Logger.log("Found matching row for notification: " + rowName + " - " + rowProject);
        
        // Kumpulkan semua previous approvers berdasarkan layer rejection
        if (layer === "SECOND_LAYER" || layer === "THIRD_LAYER") {
          if (row[10] && row[7] === "APPROVED") { // FirstLayerEmail & status approved
            previousApprovers.push({
              email: row[10],
              layer: "First Layer",
              name: "First Layer Approver"
            });
          }
        }
        if (layer === "THIRD_LAYER") {
          if (row[11] && row[8] === "APPROVED") { // SecondLayer Email & status approved
            previousApprovers.push({
              email: row[11],
              layer: "Second Layer", 
              name: "Second Layer Approver"
            });
          }
        }
        break;
      }
    }

    // Kirim email ke previous approvers
    previousApprovers.forEach(function(approver) {
      var approverSubject = "Update: Approval Request Rejected - " + project;
      var approverBody = "Dear " + (approver.name || "Approver") + ",\n\n" +
        "An approval request that you previously approved has been REJECTED at a later stage.\n\n" +
        "Project: " + project + "\n" +
        "Requester: " + name + " (" + email + ")\n" +
        "Your Approval: " + approver.layer + "\n" +
        "Rejected at: " + layerDisplay + " Approval\n" +
        "Rejection Date: " + currentTime + "\n\n" +
        "REJECTION REASON:\n" + rejectionNote + "\n\n" +
        "The requester has been notified and asked to make necessary revisions.\n\n" +
        "Best regards,\n" +
        companyName + " Approval System";

      MailApp.sendEmail({
        to: approver.email,
        subject: approverSubject,
        body: approverBody
      });

      Logger.log("Rejection notification sent to previous approver: " + approver.email);
    });

    // 3. NOTIFY ADMIN (optional - untuk tracking)
    try {
      var adminSubject = "Rejection Recorded - " + project;
      var adminBody = "Rejection recorded in Multi-Layer Approval System:\n\n" +
        "Project: " + project + "\n" +
        "Requester: " + name + " (" + email + ")\n" +
        "Rejected at: " + layerDisplay + "\n" +
        "Rejection Reason: " + rejectionNote + "\n" +
        "Time: " + currentTime + "\n\n" +
        "System: Multi-Layer Approval";
      
      sendAdminNotification(adminBody, adminSubject);
    } catch (adminError) {
      Logger.log("Admin notification optional: " + adminError.toString());
    }

    Logger.log("All rejection notifications sent successfully");

  } catch (error) {
    Logger.log("Error sending rejection notification: " + error.toString());
    throw error; // Re-throw agar tahu di log utama
  }
}

// NEW: IMPROVED REJECTION SUCCESS PAGE
// FIXED: REJECTION SUCCESS PAGE - SAMA KAYA APPROVAL
function createRejectionSuccessPage(name, email, project, documentType, attachment, layer, rejectionNote) {
  var layerDisplayNames = {
    "FIRST_LAYER": "First Layer",
    "SECOND_LAYER": "Second Layer",
    "THIRD_LAYER": "Third Layer"
  };

  var layerDisplay = layerDisplayNames[layer] || layer;
  var currentTime = getGMT7Time();

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
      max-width: 500px;
      width: 90%;
      margin: 0 auto;
    }
    .success-icon {
      font-size: 64px;
      margin-bottom: 20px;
      color: #dc2626;
    }
    .title {
      font-weight: 700;
      font-size: 28px;
      margin-bottom: 8px;
      color: #dc2626;
    }
    .details {
      background: #f8fafc;
      padding: 20px;
      border-radius: 10px;
      margin: 20px 0;
      text-align: left;
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
    .rejection-note {
      background: #fef2f2;
      padding: 15px;
      border-radius: 8px;
      margin: 20px 0;
      border-left: 4px solid #dc2626;
      text-align: left;
    }
    .close-note {
      font-size: 13px;
      color: #64748b;
      margin-top: 20px;
      font-style: italic;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="success-icon">✓</div>
    <h1 class="title">Rejection Submitted Successfully</h1>
    <p style="color: #dc2626; margin-bottom: 25px;">Thank you for your review</p>
    
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
        <span class="detail-label">Layer:</span>
        <span class="detail-value">${layerDisplay}</span>
      </div>
      <div class="detail-item">
        <span class="detail-label">Date:</span>
        <span class="detail-value">${currentTime}</span>
      </div>
    </div>
    
    <div class="rejection-note">
      <strong style="color:#dc2626;">Your Rejection Reason:</strong><br>
      ${rejectionNote.replace(/\n/g, '<br>')}
    </div>
    
    <p class="close-note">The rejection has been recorded and notifications have been sent. You can safely close this page.</p>
    
    <div style="margin-top: 25px; padding-top: 20px; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b;">
      <strong>Atreus Global</strong> • Multi-Layer Approval System
    </div>
  </div>
  
  <script>
    // Auto-close after 5 seconds
    setTimeout(function() {
      window.close();
    }, 5000);
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// PART 3: SUCCESS PAGES & MENU FUNCTIONS
// ============================================

function showRejectionSuccessPage(params) {
  try {
    var name = params.name ? decodeURIComponent(params.name) : "";
    var project = params.project ? decodeURIComponent(params.project) : "";
    var layer = params.layer ? decodeURIComponent(params.layer) : "";
    
    var layerDisplayNames = {
      "FIRST_LAYER": "First Layer",
      "SECOND_LAYER": "Second Layer", 
      "THIRD_LAYER": "Third Layer"
    };
    
    var layerDisplay = layerDisplayNames[layer] || layer;
    
    var html = `
<!DOCTYPE html>
<html>
<head>
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
      max-width: 500px;
      width: 90%;
      margin: 0 auto;
    }
    .success-icon {
      font-size: 64px;
      margin-bottom: 20px;
      color: #dc2626;
    }
    .title {
      font-weight: 700;
      font-size: 28px;
      margin-bottom: 8px;
      color: #dc2626;
    }
    .details {
      background: #f8fafc;
      padding: 20px;
      border-radius: 10px;
      margin: 20px 0;
      text-align: left;
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
    .close-note {
      font-size: 13px;
      color: #64748b;
      margin-top: 20px;
      font-style: italic;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="success-icon">✓</div>
    <h1 class="title">Rejection Submitted Successfully</h1>
    <p style="color: #dc2626; margin-bottom: 25px;">Thank you for your review</p>
    
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
        <span class="detail-label">Layer:</span>
        <span class="detail-value">${layerDisplay}</span>
      </div>
      <div class="detail-item">
        <span class="detail-label">Date:</span>
        <span class="detail-value">${getGMT7Time()}</span>
      </div>
    </div>
    
    <p class="close-note">The rejection has been recorded and notifications have been sent. You can safely close this page.</p>
    
    <div style="margin-top: 25px; padding-top: 20px; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b;">
      <strong>Atreus Global</strong> • Multi-Layer Approval System
    </div>
  </div>
  
  <script>
    // Auto-close after 5 seconds
    setTimeout(function() {
      window.close();
    }, 5000);
  </script>
</body>
</html>`;
    
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  } catch (error) {
    Logger.log("Error in showRejectionSuccessPage: " + error.toString());
    return createErrorPage("Error showing success page: " + error.message);
  }
}

function createSuccessPage(name, email, project, documentType, attachment, layer, code) {
  var layerDisplayNames = {
    "FIRST_LAYER": "First Layer",
    "SECOND_LAYER": "Second Layer",
    "THIRD_LAYER": "Third Layer"
  };
  
  var layerDisplay = layerDisplayNames[layer] || layer;
  
  var docTypeBadge = "";
  if (documentType && documentType !== "") {
    var docTypeColors = {
      "ICC": "#3B82F6",
      "Quotation": "#10B981",
      "Proposal": "#F59E0B"
    };
    var badgeColor = docTypeColors[documentType] || "#6B7280";
    docTypeBadge = '<div style="display: inline-block; background: ' + badgeColor + '; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; margin: 10px 0;">' + documentType + '</div>';
  }
  
  var attachmentSection = "";
  if (attachment && attachment !== "") {
    attachmentSection = '<div class="detail-item"><span class="detail-label">Attachment:</span><span class="detail-value"><a href="' + attachment + '" target="_blank" style="color: #326BC6;">View Document</a></span></div>';
  }
  
  var nextStep = "";
  var completionStatus = "";
  
  if (layer === "FIRST_LAYER") {
    nextStep = "This approval will now proceed to Second Layer review.";
    completionStatus = "First Layer Approval Complete";
  } else if (layer === "SECOND_LAYER") {
    nextStep = "This approval will now proceed to Third Layer final approval.";
    completionStatus = "Second Layer Approval Complete";
  } else if (layer === "THIRD_LAYER") {
    nextStep = "This project has been fully approved and completed!";
    completionStatus = "Final Approval Complete";
  }
  
  var progressBar = getProgressBarForSuccessPage(layer);
  
  var html = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%;margin:0 auto}.success-icon{font-size:64px;margin-bottom:20px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;font-weight:700}.title{font-weight:700;font-size:28px;margin-bottom:8px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}.subtitle{font-weight:500;font-size:16px;color:#326BC6;margin-bottom:25px}.completion-badge{background:linear-gradient(135deg,#d1fae5 0%,#a7f3d0 100%);color:#065f46;padding:8px 16px;border-radius:20px;font-size:14px;font-weight:600;margin-bottom:20px;display:inline-block}.details{background:linear-gradient(135deg,#f8fafc 0%,#f1f5f9 100%);padding:20px;border-radius:10px;margin:20px 0;text-align:left;border-left:4px solid #326BC6}.detail-item{margin-bottom:10px;display:flex}.detail-label{font-weight:600;color:#183460;min-width:120px}.detail-value{color:#334155;flex:1}.next-step{background:linear-gradient(135deg,#d1fae5 0%,#a7f3d0 100%);padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid #059669}.next-step strong{color:#065f46}.close-note{font-size:13px;color:#64748b;margin-top:20px;font-style:italic}.company-brand{margin-top:25px;padding-top:20px;border-top:1px solid #e2e8f0;font-size:12px;color:#64748b}.progress-track{background:#f8f9fa;padding:15px;border-radius:8px;margin:20px 0}.progress-steps{display:flex;justify-content:center;align-items:center;gap:40px;position:relative}.progress-step{text-align:center;position:relative}.step-number{width:40px;height:40px;border-radius:50%;margin:0 auto 8px;font-weight:600;display:flex;align-items:center;justify-content:center;font-size:16px}.step-active{background:#326BC6;color:white}.step-completed{background:#10b981;color:white}.step-pending{background:#e0e0e0;color:#666}.step-label{font-size:11px;font-weight:500;color:#64748b}.step-current .step-label{color:#326BC6;font-weight:600}</style></head><body><div class="container"><div class="success-icon">✓</div><h1 class="title">' + layerDisplay + ' Approval Successful!</h1><div class="completion-badge">' + completionStatus + '</div><p class="subtitle">Thank you <strong>' + (name || "User") + '</strong></p>' + docTypeBadge + '<div class="progress-track">' + progressBar + '</div><div class="details"><div class="detail-item"><span class="detail-label">Project:</span><span class="detail-value">' + (project || "N/A") + '</span></div><div class="detail-item"><span class="detail-label">Approver:</span><span class="detail-value">' + (name || "N/A") + '</span></div><div class="detail-item"><span class="detail-label">Email:</span><span class="detail-value">' + (email || "N/A") + '</span></div><div class="detail-item"><span class="detail-label">Approval Layer:</span><span class="detail-value">' + layerDisplay + '</span></div><div class="detail-item"><span class="detail-label">Date:</span><span class="detail-value">' + getGMT7Time() + '</span></div><div class="detail-item"><span class="detail-label">Approval Code:</span><span class="detail-value">' + (code || "N/A") + '</span></div>' + attachmentSection + '</div><div class="next-step"><strong>Next Step:</strong> ' + nextStep + '</div><p class="close-note">You can safely close this page.</p><div class="company-brand"><strong>Atreus Global</strong> • Multi-Layer Approval System</div></div></body></html>';
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getProgressBarForSuccessPage(currentLayer) {
  var steps = [
    { number: 1, label: "First Layer", status: currentLayer === "FIRST_LAYER" ? "active" : (["SECOND_LAYER", "THIRD_LAYER"].includes(currentLayer) ? "completed" : "pending") },
    { number: 2, label: "Second Layer", status: currentLayer === "SECOND_LAYER" ? "active" : (currentLayer === "THIRD_LAYER" ? "completed" : "pending") },
    { number: 3, label: "Third Layer", status: currentLayer === "THIRD_LAYER" ? "active" : "pending" }
  ];
  
  var progressHtml = '<div class="progress-steps">';
  
  steps.forEach(function(step) {
    var stepClass = "progress-step";
    if (step.status === "active") {
      stepClass += " step-current";
    }
    
    progressHtml += '<div class="' + stepClass + '"><div class="step-number step-' + step.status + '">' + step.number + '</div><div class="step-label">' + step.label + '</div></div>';
  });
  
  progressHtml += '</div>';
  return progressHtml;
}

function createErrorPage(errorMessage) {
  var html = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%;margin:0 auto}.error-icon{font-size:64px;margin-bottom:20px;color:#dc2626;font-weight:700}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#dc2626}.error-message{background:linear-gradient(135deg,#fef2f2 0%,#fee2e2 100%);padding:18px;border-radius:8px;margin:20px 0;border-left:4px solid #dc2626;text-align:left}.support-contact{font-size:13px;color:#64748b;margin-top:20px;padding:15px;background:#f8fafc;border-radius:8px}.action-buttons{margin-top:25px;display:flex;gap:10px;justify-content:center}.btn{padding:10px 20px;border:none;border-radius:6px;font-weight:500;cursor:pointer;text-decoration:none;display:inline-block;transition:all 0.2s ease}.btn-primary{background:#326BC6;color:white}.btn-secondary{background:#64748b;color:white}.btn:hover{transform:translateY(-1px);box-shadow:0 4px 8px rgba(0,0,0,0.1)}</style></head><body><div class="container"><div class="error-icon">✗</div><h1 class="title">Approval Failed</h1><div class="error-message"><strong style="color:#dc2626;">Error Details:</strong><br>' + errorMessage + '</div><p style="color:#475569;margin-bottom:20px;">Please try again or contact your system administrator if this issue persists.</p><div class="action-buttons"><a href="javascript:window.location.reload()" class="btn btn-primary">Try Again</a><a href="javascript:window.close()" class="btn btn-secondary">Close</a></div><div class="support-contact"><strong style="color:#183460;">Atreus Global</strong><br>Please provide the error message above when contacting support.<br><span style="font-size:11px;color:#94a3b8;">Error ID: ' + Utilities.getUuid().substring(0, 8) + '</span></div></div></body></html>';
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// MENU & UTILITY FUNCTIONS
// ============================================

function showApprovalPipeline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
  
  var pipeline = {
    "PENDING_FIRST_LAYER": [],
    "PENDING_SECOND_LAYER": [],
    "PENDING_THIRD_LAYER": [],
    "COMPLETED": [],
    "INVALID_ATTACHMENT": []
  };

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var firstLayer = row[7]; // Column H
    var secondLayer = row[8]; // Column I
    var thirdLayer = row[9]; // Column J
    var attachment = row[4]; // Column E
    var documentType = row[3]; // Column D

    if (attachment && attachment !== "") {
      var validation = validateGoogleDriveAttachmentWithType(attachment, documentType);
      if (!validation.valid) {
        pipeline.INVALID_ATTACHMENT.push(row[0] + " - " + row[2]);
        continue;
      }
    }

    if (!firstLayer || firstLayer === "PENDING") {
      pipeline.PENDING_FIRST_LAYER.push(row[0] + " - " + row[2]);
    } else if (firstLayer === "APPROVED" && (!secondLayer || secondLayer === "PENDING")) {
      pipeline.PENDING_SECOND_LAYER.push(row[0] + " - " + row[2]);
    } else if (secondLayer === "APPROVED" && (!thirdLayer || thirdLayer === "PENDING")) {
      pipeline.PENDING_THIRD_LAYER.push(row[0] + " - " + row[2]);
    } else if (thirdLayer === "APPROVED") {
      pipeline.COMPLETED.push(row[0] + " - " + row[2]);
    }
  }

  var message = "MULTI-LAYER APPROVAL PIPELINE\n\n";
  message += "Pending First Layer: " + pipeline.PENDING_FIRST_LAYER.length + "\n";
  message += "Pending Second Layer: " + pipeline.PENDING_SECOND_LAYER.length + "\n";
  message += "Pending Third Layer: " + pipeline.PENDING_THIRD_LAYER.length + "\n";
  message += "Completed: " + pipeline.COMPLETED.length + "\n";
  message += "Invalid Attachment: " + pipeline.INVALID_ATTACHMENT.length + "\n\n";
  
  if (pipeline.INVALID_ATTACHMENT.length > 0) {
    message += "Invalid Attachments:\n";
    pipeline.INVALID_ATTACHMENT.forEach(function(item) {
      message += "- " + item + "\n";
    });
  }
  
  SpreadsheetApp.getUi().alert("Approval Pipeline Dashboard", message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function validateAllAttachments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
  
  var validationResults = {
    valid: 0,
    invalid: 0,
    empty: 0,
    sharedDrive: 0,
    myDrive: 0,
    details: []
  };
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var attachment = row[4]; // Column E
    var documentType = row[3]; // Column D
    var validation = validateGoogleDriveAttachmentWithType(attachment, documentType);
    
    validationResults.details.push({
      row: i + 2,
      project: row[2],
      documentType: documentType,
      attachment: attachment,
      validation: validation
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
  
  var message = "ATTACHMENT VALIDATION REPORT\n\n";
  message += "Valid: " + validationResults.valid + "\n";
  message += "  - Shared Drive: " + validationResults.sharedDrive + "\n";
  message += "  - My Drive: " + validationResults.myDrive + "\n";
  message += "Invalid: " + validationResults.invalid + "\n";
  message += "Empty: " + validationResults.empty + "\n\n";
  
  if (validationResults.invalid > 0) {
    message += "Invalid Attachments Found:\n";
    validationResults.details.forEach(function(detail) {
      if (!detail.validation.valid && detail.attachment) {
        message += "- Row " + detail.row + ": " + detail.project + " - " + detail.validation.message + "\n";
      }
    });
  }
  
  SpreadsheetApp.getUi().alert("Attachment Validation", message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function resetMultiLayerRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = sheet.getActiveRange();
  var startRow = activeRange.getRow();
  
  for (var i = startRow; i < startRow + activeRange.getNumRows(); i++) {
    if (i >= 2) {
      // Reset checkboxes and status
      sheet.getRange(i, 6).setValue(false); // Column F - Checkbox
      sheet.getRange(i, 7).setNote(""); // Column G - Log
      sheet.getRange(i, 8).setValue("PENDING"); // Column H - FirstLayer
      sheet.getRange(i, 8).setBackground(null);
      sheet.getRange(i, 9).setValue("PENDING"); // Column I - Second Layer
      sheet.getRange(i, 9).setBackground(null);
      sheet.getRange(i, 10).setValue("PENDING"); // Column J - ThirdLayer
      sheet.getRange(i, 10).setBackground(null);
      sheet.getRange(i, 14).setValue("ACTIVE"); // Column N - Status
      sheet.getRange(i, 14).setBackground(null);
    }
  }
  
  SpreadsheetApp.getUi().alert("Reset Complete", "Selected multi-layer rows have been reset!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function testCompleteFlow() {
  Logger.log("Testing Complete Approval Flow...");

  var testName = "Test User";
  var testEmail = "test@atreusg.com";
  var testProject = "Test Project";
  var testLayer = "FIRST_LAYER";
  var testCode = "test123";

  var result = updateMultiLayerApprovalStatus(testName, testEmail, testProject, testLayer, testCode);

  if (result) {
    Logger.log(" Test PASSED");
    SpreadsheetApp.getUi().alert("Test Result", " Test PASSED - Approval workflow is working!", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    Logger.log("Test FAILED");
    SpreadsheetApp.getUi().alert("Test Result", "Test FAILED - Check logs for details.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function manualApproveRow(rowNumber, layer) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rowData = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];

    var name = rowData[0];
    var email = rowData[1];
    var project = rowData[2];
    var code = "manual_" + new Date().getTime();

    var result = updateMultiLayerApprovalStatus(name, email, project, layer, code);

    if (result) {
      SpreadsheetApp.getUi().alert("Success", "Manual approval completed for " + name + " - Layer: " + layer, SpreadsheetApp.getUi().ButtonSet.OK);

      if (layer === "FIRST_LAYER") {
        sendNextApprovalAfterFirstLayer();
      } else if (layer === "SECOND_LAYER") {
        sendNextApprovalAfterSecondLayer();
      }
    } else {
      SpreadsheetApp.getUi().alert("Failed", "Manual approval failed for " + name + ". Check logs.", SpreadsheetApp.getUi().ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert("Error", "Manual approval error: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function forceSendNextLayers() {
  try {
    sendNextApprovalAfterFirstLayer();
    Utilities.sleep(2000);
    sendNextApprovalAfterSecondLayer();
    SpreadsheetApp.getUi().alert("Force Send Complete", "Checked and sent all pending next layer approvals.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error", "Force send error: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function viewLogs() {
  SpreadsheetApp.getUi().alert("View Logs", "To view logs:\n1. Go to Extensions > Apps Script\n2. Click 'Executions' on left sidebar\n3. See recent runs and logs.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function manualApproveFirstLayer() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Invalid Row", "Please select a row starting from row 2.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  manualApproveRow(row, "FIRST_LAYER");
}

function manualApproveSecondLayer() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Invalid Row", "Please select a row starting from row 2.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  manualApproveRow(row, "SECOND_LAYER");
}

function manualApproveThirdLayer() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Invalid Row", "Please select a row starting from row 2.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  manualApproveRow(row, "THIRD_LAYER");
}

function testRejectionFlow() {
  Logger.log("Testing Rejection Flow for All Layers...");
  
  var testData = [
    { layer: "FIRST_LAYER", name: "Test User 1", project: "Test Project First Layer" },
    { layer: "SECOND_LAYER", name: "Test User 2", project: "Test Project Second Layer" },
    { layer: "THIRD_LAYER", name: "Test User 3", project: "Test Project Third Layer" }
  ];
  
  testData.forEach(function(test) {
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
      Logger.log("" + test.layer + " rejection: PASSED");
    } else {
      Logger.log("" + test.layer + " rejection: FAILED");
    }
  });
  
  SpreadsheetApp.getUi().alert(
    "Rejection Test Complete", 
    "Check logs for results. Make sure to check:\n" +
    "1. Status changed to REJECTED\n" + 
    "2. Subsequent layers reset to PENDING\n" +
    "3. Background colors updated", 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Multi-Layer Approval')
    .addItem('Send Multi-Layer Approvals', 'sendMultiLayerApproval')
    .addItem('Send Second Layer Approval (After First)', 'sendNextApprovalAfterFirstLayer')
    .addItem('Send Third Layer Approval (After Second)', 'sendNextApprovalAfterSecondLayer')
    .addItem('Force Send Next Layer (All Pending)', 'forceSendNextLayers')
    .addSeparator()
    .addItem('View Approval Pipeline', 'showApprovalPipeline')
    .addItem('Check Attachment Validation', 'validateAllAttachments')
    .addItem('Reset Selected Rows', 'resetMultiLayerRows')
    .addSeparator()
    .addItem('Test Complete Flow', 'testCompleteFlow')
    .addItem('Manual Approve Selected Row - First Layer', 'manualApproveFirstLayer')
    .addItem('Manual Approve Selected Row - Second Layer', 'manualApproveSecondLayer')
    .addItem('Manual Approve Selected Row - Third Layer', 'manualApproveThirdLayer')
    .addSeparator()
    .addItem('Test Rejection Flow (All Layers)', 'testRejectionFlow')
    .addItem('Debug Current Row', 'debugCurrentRow')
    .addSeparator()
    .addItem('View Recent Logs', 'viewLogs')
    .addToUi();
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

function sendTestEmail() {
  try {
    var testEmail = "test@atreusg.com";
    var testDescription = "Test Project";
    var testDocumentType = "ICC";
    var testAttachment = "https://drive.google.com/file/d/1abc123/view";
    var testApprovalLink = "https://script.google.com/macros/s/AKfycbwwUmgrdFtYQjx_EuXStjAXy3RPhExlsew1VZc5ZKeWDSlh96c-sm9dkYkL_3-k2Tvf/exec?action=approve&layer=FIRST_LAYER";
    var testValidation = {
      valid: true,
      message: "Valid Google Drive file",
      name: "test_document.pdf",
      type: "application/pdf",
      sizeFormatted: "2.5 MB",
      isSharedDrive: false
    };
    
    var result = sendMultiLayerEmail(testEmail, testDescription, testDocumentType, testAttachment, testApprovalLink, "FIRST_LAYER", testValidation);
    
    if (result) {
      Logger.log(" Test email sent successfully");
      SpreadsheetApp.getUi().alert("Test Email", "Test email sent successfully to " + testEmail, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Test email failed");
      SpreadsheetApp.getUi().alert("Test Email", "Test email failed to send", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log("Test email error: " + error.toString());
    SpreadsheetApp.getUi().alert("Test Email Error", "Error: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function debugCurrentRow() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = sheet.getActiveRange().getRow();
    
    if (row < 2) {
      SpreadsheetApp.getUi().alert("Invalid Row", "Please select a row starting from row 2.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var rowData = sheet.getRange(row, 1, 1, 14).getValues()[0];
    
    var debugInfo = "DEBUG INFO - Row " + row + "\n\n";
    debugInfo += "Name: " + rowData[0] + "\n";
    debugInfo += "Email: " + rowData[1] + "\n";
    debugInfo += "Description: " + rowData[2] + "\n";
    debugInfo += "Document Type: " + rowData[3] + "\n";
    debugInfo += "Attachment: " + rowData[4] + "\n";
    debugInfo += "Send Status: " + rowData[5] + "\n";
    debugInfo += "Log: " + rowData[6] + "\n";
    debugInfo += "First Layer Status: " + rowData[7] + "\n";
    debugInfo += "Second Layer Status: " + rowData[8] + "\n";
    debugInfo += "Third Layer Status: " + rowData[9] + "\n";
    debugInfo += "First Layer Email: " + rowData[10] + "\n";
    debugInfo += "Second Layer Email: " + rowData[11] + "\n";
    debugInfo += "Third Layer Email: " + rowData[12] + "\n";
    debugInfo += "Overall Status: " + rowData[13] + "\n";
    
    // Validate attachment
    if (rowData[4]) {
      var validation = validateGoogleDriveAttachmentWithType(rowData[4], rowData[3]);
      debugInfo += "\nATTACHMENT VALIDATION:\n";
      debugInfo += "Valid: " + validation.valid + "\n";
      debugInfo += "Message: " + validation.message + "\n";
      debugInfo += "File Name: " + (validation.name || "N/A") + "\n";
      debugInfo += "File Type: " + (validation.type || "N/A") + "\n";
      debugInfo += "Drive Type: " + (validation.driveType || "N/A") + "\n";
    }
    
    SpreadsheetApp.getUi().alert("Debug Information", debugInfo, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("Debug Error", "Error: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}