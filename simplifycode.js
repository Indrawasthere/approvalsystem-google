// ============================================
// FULL MULTI-LAYER APPROVAL SYSTEM
// Copy-paste langsung ke Apps Script
// ============================================

var DEPLOYMENT_ID = "AKfycbz4ezLlNBndJqw334RXIQbo3ojoEcX9E3eazRjqJbUH7YYK0OYdwSzLKAmhwI2GramZ";
var WEB_APP_URL = "https://script.google.com/macros/s/" + DEPLOYMENT_ID + "/exec";

// ============================================
// MAIN FUNCTION - SEND APPROVALS
// ============================================

function sendMultiLayerApproval() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("No data to process!");
      return;
    }
    
    var data = sheet.getRange("A2:S" + lastRow).getValues();
    var results = [];
    var processedCount = 0;
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var name = row[0];
      var email = row[1];
      var description = row[2];
      var documentType = row[3];
      var attachment = row[4];
      var sendStatus = row[5];
      var currentStatus = row[13];
      var currentEditor = row[14];
      var needsRevisionFrom = row[16];
      
      if (!name || !email) continue;
      
      var isChecked = (sendStatus === true || sendStatus === "TRUE" || sendStatus === "true");
      var isActive = (currentStatus === "ACTIVE" || currentStatus === "" || !currentStatus);
      var isNeedsRevision = (currentStatus === "NEEDS_REVISION" && currentEditor);
      
      var canSend = false;
      
      if (isChecked && isActive) {
        canSend = true;
      } else if (isChecked && isNeedsRevision && !currentEditor) {
        canSend = true;
      }
      
      if (!canSend) continue;
      
      Logger.log("Processing: " + name + " - " + description);
      
      var validationResult = validateGoogleDriveAttachment(attachment, documentType);
      
      var firstLayerStatus = row[7];
      var secondLayerStatus = row[8];
      var thirdLayerStatus = row[9];
      var firstLayerEmail = row[10];
      var secondLayerEmail = row[11];
      var thirdLayerEmail = row[12];
      
      var nextApproval = getNextApprovalLayer(
        firstLayerStatus, secondLayerStatus, thirdLayerStatus,
        firstLayerEmail, secondLayerEmail, thirdLayerEmail,
        currentStatus, needsRevisionFrom
      );
      
      Logger.log("Next: " + nextApproval.layer + " -> " + nextApproval.email);
      
      if (nextApproval.layer !== "COMPLETED" && nextApproval.email) {
        var approvalLink = generateApprovalLink(name, email, description, documentType, attachment, nextApproval.layer);
        var emailSent = sendApprovalEmail(nextApproval.email, description, documentType, attachment, approvalLink, nextApproval.layer, validationResult);
        
        if (emailSent) {
          sheet.getRange(i + 2, 14).setValue("PROCESSING");
          sheet.getRange(i + 2, 14).setBackground("#FFF2CC");
          sheet.getRange(i + 2, 7).setNote("Sent to " + nextApproval.layer + " - " + getGMT7Time());
          
          results.push({
            name: name,
            project: description,
            layer: nextApproval.layer,
            email: nextApproval.email
          });
          
          processedCount++;
        }
      }
      
      sheet.getRange(i + 2, 6).setValue(false);
      Utilities.sleep(500);
    }
    
    showSummary(results, processedCount);
    
  } catch (error) {
    Logger.log("Error: " + error.toString());
    SpreadsheetApp.getUi().alert("Error", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getNextApprovalLayer(firstLayer, secondLayer, thirdLayer, firstEmail, secondEmail, thirdEmail, currentStatus, needsRevisionFrom) {
  Logger.log("Current Status: " + currentStatus + ", Needs Revision From: " + needsRevisionFrom);
  
  // Jika NEEDS_REVISION, kirim balik ke layer yang reject
  if (currentStatus === "NEEDS_REVISION" && needsRevisionFrom) {
    if (needsRevisionFrom === "FIRST_LAYER") {
      return { layer: "FIRST_LAYER", email: firstEmail };
    } else if (needsRevisionFrom === "SECOND_LAYER") {
      return { layer: "SECOND_LAYER", email: secondEmail };
    } else if (needsRevisionFrom === "THIRD_LAYER") {
      return { layer: "THIRD_LAYER", email: thirdEmail };
    }
  }
  
  // Normal flow
  if (!firstLayer || firstLayer === "PENDING") {
    return { layer: "FIRST_LAYER", email: firstEmail };
  }
  if (firstLayer === "APPROVED" && (!secondLayer || secondLayer === "PENDING")) {
    return { layer: "SECOND_LAYER", email: secondEmail };
  }
  if (secondLayer === "APPROVED" && (!thirdLayer || thirdLayer === "PENDING")) {
    return { layer: "THIRD_LAYER", email: thirdEmail };
  }
  
  return { layer: "COMPLETED", email: "" };
}

// ============================================
// VALIDATION & LINK GENERATION
// ============================================

function validateGoogleDriveAttachment(attachmentUrl, documentType) {
  try {
    if (!attachmentUrl || attachmentUrl === "") {
      return { valid: false, message: "No attachment", name: "N/A", type: "N/A", sizeFormatted: "N/A", isSharedDrive: false };
    }
    
    if (!attachmentUrl.includes('drive.google.com')) {
      return { valid: false, message: "Not a Google Drive link", name: "N/A", type: "N/A", sizeFormatted: "N/A", isSharedDrive: false };
    }
    
    var fileId = extractFileId(attachmentUrl);
    if (!fileId) {
      return { valid: false, message: "Invalid URL format", name: "N/A", type: "N/A", sizeFormatted: "N/A", isSharedDrive: false };
    }
    
    try {
      var file = DriveApp.getFileById(fileId);
      
      if (file.getMimeType() === 'application/vnd.google-apps.folder') {
        return { valid: false, message: "Folder detected", name: "N/A", type: "N/A", sizeFormatted: "N/A", isSharedDrive: false };
      }
      
      var isSharedDrive = false;
      try {
        var driveFile = Drive.Files.get(fileId, { supportsAllDrives: true, fields: 'driveId' });
        isSharedDrive = driveFile.driveId != null;
      } catch (e) {
        isSharedDrive = false;
      }
      
      return {
        valid: true,
        message: "Valid file",
        name: file.getName(),
        type: file.getMimeType(),
        sizeFormatted: formatFileSize(file.getSize()),
        isSharedDrive: isSharedDrive
      };
      
    } catch (e) {
      return { valid: false, message: "File not accessible", name: "N/A", type: "N/A", sizeFormatted: "N/A", isSharedDrive: false };
    }
    
  } catch (error) {
    return { valid: false, message: "Validation error", name: "N/A", type: "N/A", sizeFormatted: "N/A", isSharedDrive: false };
  }
}

function extractFileId(url) {
  var patterns = [/\/d\/([a-zA-Z0-9_-]+)/, /id=([a-zA-Z0-9_-]+)/, /\/file\/d\/([a-zA-Z0-9_-]+)/, /\/open\?id=([a-zA-Z0-9_-]+)/];
  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match && match[1]) return match[1];
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

function generateApprovalLink(name, email, description, documentType, attachment, layer) {
  var timestamp = new Date().getTime();
  var code = generateCode(name, description, timestamp);
  
  var params = {
    action: "approve",
    name: encodeURIComponent(name || ""),
    email: encodeURIComponent(email || ""),
    project: encodeURIComponent(description || ""),
    docType: encodeURIComponent(documentType || ""),
    attachment: encodeURIComponent(attachment || ""),
    layer: layer,
    code: code,
    timestamp: timestamp
  };
  
  var query = Object.keys(params).map(function(k) { return k + "=" + params[k]; }).join("&");
  return WEB_APP_URL + "?" + query;
}

function generateCode(name, project, timestamp) {
  var nameCode = name ? name.substring(0, 3).toLowerCase() : "usr";
  var projectCode = project ? project.replace(/[^a-zA-Z0-9]/g, "").substring(0, 3).toLowerCase() : "prj";
  var timeCode = timestamp.toString(36);
  return nameCode + projectCode + timeCode;
}

function getGMT7Time() {
  var now = new Date();
  return Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
}

// ============================================
// EMAIL SENDING
// ============================================

function sendApprovalEmail(recipientEmail, description, documentType, attachment, approvalLink, layer, validationResult) {
  try {
    if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
      Logger.log("Invalid email: " + recipientEmail);
      return false;
    }
    
    var layerNames = { "FIRST_LAYER": "First Layer", "SECOND_LAYER": "Second Layer", "THIRD_LAYER": "Third Layer" };
    var layerDisplay = layerNames[layer] || layer;
    var subject = "Approval Request - " + description + " [" + layerDisplay + "]";
    
    var docBadge = documentType ? '<div style="display:inline-block;background:#3B82F6;color:white;padding:4px 12px;border-radius:12px;font-size:12px;font-weight:600;margin-bottom:10px;">' + documentType + '</div>' : '';
    
    var attachmentSec = "";
    if (attachment && attachment !== "") {
      var fileInfo = '<p><strong>File:</strong> ' + validationResult.name + '</p><p><strong>Size:</strong> ' + validationResult.sizeFormatted + '</p>';
      attachmentSec = '<div style="background:#f0f8ff;padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid #90EE90;"><h3>Attachment</h3>' + docBadge + fileInfo + '<p><a href="' + attachment + '" target="_blank" style="color:#326BC6;">View Document</a></p></div>';
    }
    
    var rejectLink = approvalLink.replace("action=approve", "action=reject");
    
    var html = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;line-height:1.6;color:#333;background:#f6f9fc;margin:0;padding:0}.container{max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1)}.header{background:linear-gradient(135deg,#326BC6 0%,#183460 100%);padding:30px;text-align:center;color:white}.content{padding:30px}.button{display:inline-block;padding:15px 30px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white !important;text-decoration:none;border-radius:8px;margin:10px;font-weight:600;border:none;cursor:pointer}.button-reject{background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white !important;text-decoration:none;border-radius:8px;margin:10px;font-weight:600;border:none;cursor:pointer}.button-container{text-align:center}.footer{margin-top:30px;padding:20px;background:#f8f9fa;text-align:center;font-size:12px;color:#666}</style></head><body><div class="container"><div class="header"><h1>Approval Required</h1><p>' + layerDisplay + '</p></div><div class="content"><h2>' + description + '</h2>' + attachmentSec + '<div class="button-container"><a href="' + approvalLink + '" class="button">APPROVE</a><a href="' + rejectLink + '" class="button-reject">REJECT</a></div></div><div class="footer"><p>© Atreus Global - Approval System</p></div></div></body></html>';
    
    MailApp.sendEmail({ to: recipientEmail, subject: subject, htmlBody: html });
    Logger.log("Email sent to: " + recipientEmail);
    return true;
    
  } catch (error) {
    Logger.log("Error sending email: " + error.toString());
    return false;
  }
}

// ============================================
// WEB APP - doGet
// ============================================

function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action;
    
    if (action === "approve") {
      return handleApproval(params);
    } else if (action === "reject") {
      return handleRejection(params);
    }
    
    return createErrorPage("Invalid request");
    
  } catch (error) {
    Logger.log("Error in doGet: " + error.toString());
    return createErrorPage("System error");
  }
}

function handleApproval(params) {
  try {
    var name = params.name ? decodeURIComponent(params.name) : "";
    var email = params.email ? decodeURIComponent(params.email) : "";
    var project = params.project ? decodeURIComponent(params.project) : "";
    var layer = params.layer ? decodeURIComponent(params.layer) : "";
    var timestamp = parseInt(params.timestamp) || 0;
    
    var now = new Date().getTime();
    var sevenDaysAgo = now - (7 * 24 * 60 * 60 * 1000);
    if (timestamp < sevenDaysAgo) {
      return createErrorPage("Link expired");
    }
    
    var result = updateApprovalStatus(name, email, project, layer, "APPROVED");
    
    if (result) {
      Utilities.sleep(2000);
      autoTriggerNextLayer();
      return createSuccessPage(name, project, layer);
    } else {
      return createErrorPage("Data not found");
    }
    
  } catch (error) {
    Logger.log("Error in handleApproval: " + error.toString());
    return createErrorPage("System error");
  }
}

function handleRejection(params) {
  try {
    var name = params.name ? decodeURIComponent(params.name) : "";
    var email = params.email ? decodeURIComponent(params.email) : "";
    var project = params.project ? decodeURIComponent(params.project) : "";
    var layer = params.layer ? decodeURIComponent(params.layer) : "";
    
    return createRejectionForm(name, email, project, layer);
    
  } catch (error) {
    Logger.log("Error in handleRejection: " + error.toString());
    return createErrorPage("System error");
  }
}

// ============================================
// WEB APP - doPost (Rejection Submission)
// ============================================

function doPost(e) {
  try {
    var params = {};
    if (e.postData && e.postData.contents) {
      var parts = e.postData.contents.split('&');
      for (var i = 0; i < parts.length; i++) {
        var pair = parts[i].split('=');
        if (pair.length === 2) {
          params[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1]);
        }
      }
    }
    
    var name = params.name || "";
    var email = params.email || "";
    var project = params.project || "";
    var layer = params.layer || "";
    var rejectionNote = params.rejectionNote || "";
    
    if (!rejectionNote || rejectionNote.trim() === "") {
      return createErrorPage("Reason required");
    }
    
    var result = updateRejectionStatus(name, email, project, layer, rejectionNote);
    
    if (result) {
      return createRejectionSuccessPage(name, project, layer);
    } else {
      return createErrorPage("Data not found");
    }
    
  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    return createErrorPage("System error");
  }
}

// ============================================
// UPDATE STATUS FUNCTIONS
// ============================================

function updateApprovalStatus(name, email, project, layer, action) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:S" + sheet.getLastRow()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var nameMatch = row[0] && row[0].toString().trim().toLowerCase() === (name || "").toString().trim().toLowerCase();
      var emailMatch = row[1] && row[1].toString().trim().toLowerCase() === (email || "").toString().trim().toLowerCase();
      var projectMatch = row[2] && row[2].toString().trim().toLowerCase() === (project || "").toString().trim().toLowerCase();
      
      if (nameMatch && emailMatch && projectMatch) {
        Logger.log("Match found at row " + (i + 2));
        
        var layerCol = getLayerColumn(layer);
        if (layerCol === -1) return false;
        
        var currentStatus = sheet.getRange(i + 2, layerCol).getValue();
        if (currentStatus === "APPROVED" || currentStatus === "REJECTED") {
          return false;
        }
        
        // Update approval status
        sheet.getRange(i + 2, layerCol).setValue("APPROVED");
        sheet.getRange(i + 2, layerCol).setBackground("#90EE90");
        
        // Clear editing fields
        sheet.getRange(i + 2, 15).setValue("");
        sheet.getRange(i + 2, 16).setValue("");
        sheet.getRange(i + 2, 17).setValue("");
        
        // Check if all approved
        var col1 = sheet.getRange(i + 2, 8).getValue();
        var col2 = sheet.getRange(i + 2, 9).getValue();
        var col3 = sheet.getRange(i + 2, 10).getValue();
        
        if (col1 === "APPROVED" && col2 === "APPROVED" && col3 === "APPROVED") {
          sheet.getRange(i + 2, 14).setValue("COMPLETED");
          sheet.getRange(i + 2, 14).setBackground("#90EE90");
        } else {
          sheet.getRange(i + 2, 14).setValue("PROCESSING");
          sheet.getRange(i + 2, 14).setBackground("#FFF2CC");
        }
        
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return false;
  }
}

function updateRejectionStatus(name, email, project, layer, rejectionNote) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:S" + sheet.getLastRow()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var nameMatch = row[0] && row[0].toString().trim().toLowerCase() === (name || "").toString().trim().toLowerCase();
      var emailMatch = row[1] && row[1].toString().trim().toLowerCase() === (email || "").toString().trim().toLowerCase();
      var projectMatch = row[2] && row[2].toString().trim().toLowerCase() === (project || "").toString().trim().toLowerCase();
      
      if (nameMatch && emailMatch && projectMatch) {
        Logger.log("Match found for rejection at row " + (i + 2));
        
        var layerCol = getLayerColumn(layer);
        if (layerCol === -1) return false;
        
        var currentStatus = sheet.getRange(i + 2, layerCol).getValue();
        if (currentStatus === "REJECTED" || currentStatus === "APPROVED") {
          return false;
        }
        
        // Determine editor based on rejection layer
        var editorEmail = "";
        if (layer === "FIRST_LAYER") {
          editorEmail = row[1];
        } else if (layer === "SECOND_LAYER") {
          editorEmail = row[10];
        } else if (layer === "THIRD_LAYER") {
          editorEmail = row[11];
        }
        
        // Reset subsequent layers
        if (layer === "FIRST_LAYER") {
          sheet.getRange(i + 2, 9).setValue("PENDING");
          sheet.getRange(i + 2, 10).setValue("PENDING");
        } else if (layer === "SECOND_LAYER") {
          sheet.getRange(i + 2, 10).setValue("PENDING");
        }
        
        // Update rejection status
        sheet.getRange(i + 2, layerCol).setValue("REJECTED");
        sheet.getRange(i + 2, layerCol).setBackground("#FF6B6B");
        
        // Set editing responsibility
        sheet.getRange(i + 2, 15).setValue(editorEmail);
        sheet.getRange(i + 2, 16).setValue(rejectionNote);
        sheet.getRange(i + 2, 17).setValue(layer);
        
        // Update overall status
        sheet.getRange(i + 2, 14).setValue("NEEDS_REVISION");
        sheet.getRange(i + 2, 14).setBackground("#FFE4CC");
        
        // Send notifications
        sendRejectionNotifications(row[0], row[1], project, layer, rejectionNote, editorEmail);
        
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return false;
  }
}

function getLayerColumn(layer) {
  if (layer === "FIRST_LAYER") return 8;
  if (layer === "SECOND_LAYER") return 9;
  if (layer === "THIRD_LAYER") return 10;
  return -1;
}

// ============================================
// AUTO-TRIGGER NEXT LAYER
// ============================================

function autoTriggerNextLayer() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange("A2:S" + sheet.getLastRow()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var status = row[13];
      
      if (status !== "PROCESSING") continue;
      
      var firstStatus = row[7];
      var secondStatus = row[8];
      var thirdStatus = row[9];
      
      // First layer approved, send to second
      if (firstStatus === "APPROVED" && (!secondStatus || secondStatus === "PENDING")) {
        var secondEmail = row[11];
        if (secondEmail) {
          var validation = validateGoogleDriveAttachment(row[4], row[3]);
          var link = generateApprovalLink(row[0], row[1], row[2], row[3], row[4], "SECOND_LAYER");
          sendApprovalEmail(secondEmail, row[2], row[3], row[4], link, "SECOND_LAYER", validation);
          Utilities.sleep(500);
        }
      }
      
      // Second layer approved, send to third
      if (secondStatus === "APPROVED" && (!thirdStatus || thirdStatus === "PENDING")) {
        var thirdEmail = row[12];
        if (thirdEmail) {
          var validation = validateGoogleDriveAttachment(row[4], row[3]);
          var link = generateApprovalLink(row[0], row[1], row[2], row[3], row[4], "THIRD_LAYER");
          sendApprovalEmail(thirdEmail, row[2], row[3], row[4], link, "THIRD_LAYER", validation);
          Utilities.sleep(500);
        }
      }
    }
  } catch (error) {
    Logger.log("Error auto-triggering: " + error.toString());
  }
}

// ============================================
// NOTIFICATIONS
// ============================================

function sendRejectionNotifications(requesterName, requesterEmail, project, layer, rejectionNote, editorEmail) {
  try {
    var layerNames = { "FIRST_LAYER": "First Layer", "SECOND_LAYER": "Second Layer", "THIRD_LAYER": "Third Layer" };
    var layerDisplay = layerNames[layer] || layer;
    
    var requesterSubject = "Feedback on Submission - " + project;
    var requesterBody = "Dear " + requesterName + ",\n\n" +
      "Your submission has been reviewed by " + layerDisplay + " and requires revision.\n\n" +
      "Feedback:\n" + rejectionNote + "\n\n" +
      "Please make the necessary changes and resubmit.\n\n" +
      "Best regards,\nApproval System";
    
    MailApp.sendEmail({ to: requesterEmail, subject: requesterSubject, body: requesterBody });
    
    if (editorEmail && editorEmail !== requesterEmail) {
      var editorSubject = "Action Required: Edit Document - " + project;
      var editorBody = "Dear Editor,\n\n" +
        "Document '" + project + "' rejected at " + layerDisplay + ".\n\n" +
        "Feedback:\n" + rejectionNote + "\n\n" +
        "Requester: " + requesterName + " (" + requesterEmail + ")\n\n" +
        "Best regards,\nApproval System";
      
      MailApp.sendEmail({ to: editorEmail, subject: editorSubject, body: editorBody });
    }
  } catch (error) {
    Logger.log("Error sending notifications: " + error.toString());
  }
}

// ============================================
// HTML PAGES
// ============================================

function createRejectionForm(name, email, project, layer) {
  var layerNames = { "FIRST_LAYER": "First Layer", "SECOND_LAYER": "Second Layer", "THIRD_LAYER": "Third Layer" };
  var layerDisplay = layerNames[layer] || layer;
  
  var html = '<!DOCTYPE html><html><head><base target="_top"><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;padding:20px;background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:600px;width:90%}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#dc2626}.details{background:#f8fafc;padding:20px;border-radius:10px;margin:20px 0;border-left:4px solid #dc2626}.detail-item{margin-bottom:10px;display:flex}.detail-label{font-weight:600;color:#dc2626;min-width:100px}.detail-value{color:#334155;flex:1}.form-group{margin-bottom:15px}.form-label{font-weight:600;color:#dc2626;margin-bottom:5px;display:block}.form-textarea{width:100%;padding:10px;border:1px solid #e5e7eb;border-radius:5px;font-family:inherit;min-height:120px;box-sizing:border-box}.button{padding:15px 30px;background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white !important;border-radius:8px;font-weight:600;border:none;cursor:pointer;transition:all 0.3s}.button:hover{transform:translateY(-2px);box-shadow:0 6px 12px rgba(220,38,38,0.3)}.button:disabled{opacity:0.6}.loading{display:none;color:#dc2626;font-weight:600;margin-top:15px}</style></head><body><div class="container"><h1 class="title">Rejection Form</h1><div class="details"><div class="detail-item"><span class="detail-label">Project:</span><span class="detail-value">' + project + '</span></div><div class="detail-item"><span class="detail-label">Layer:</span><span class="detail-value">' + layerDisplay + '</span></div><div class="detail-item"><span class="detail-label">Date:</span><span class="detail-value">' + getGMT7Time() + '</span></div></div><form id="rejectionForm" onsubmit="return false;"><input type="hidden" name="name" value="' + (name || '') + '"><input type="hidden" name="email" value="' + (email || '') + '"><input type="hidden" name="project" value="' + (project || '') + '"><input type="hidden" name="layer" value="' + (layer || '') + '"><div class="form-group"><label class="form-label">Rejection Reason (Required):</label><textarea id="rejectionNote" name="rejectionNote" class="form-textarea" placeholder="Explain what needs to be improved..." required></textarea></div><button type="button" class="button" onclick="submitRejection()">Submit Rejection</button><div id="loading" class="loading">Processing... Please wait.</div></form></div><script>function submitRejection(){var note=document.getElementById("rejectionNote").value;if(!note.trim()){alert("Please provide a reason");return;}document.querySelector(".button").disabled=true;document.getElementById("loading").style.display="block";var form=document.getElementById("rejectionForm");var data=new FormData(form);var params={};for(var p of data.entries()){params[p[0]]=p[1];}params.rejectionNote=note;google.script.run.withSuccessHandler(function(){alert("Submitted!");setTimeout(function(){window.close();},1000);}).withFailureHandler(function(e){alert("Error: "+e);document.querySelector(".button").disabled=false;document.getElementById("loading").style.display="none";}).handleRejectionFromWeb(params);}</script></body></html>';
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createSuccessPage(name, project, layer) {
  var layerNames = { "FIRST_LAYER": "First Layer", "SECOND_LAYER": "Second Layer", "THIRD_LAYER": "Third Layer" };
  var layerDisplay = layerNames[layer] || layer;
  
  var html = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%}.icon{font-size:64px;margin-bottom:20px}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#326BC6}.details{background:#f8fafc;padding:20px;border-radius:10px;margin:20px 0;text-align:left;border-left:4px solid #326BC6}.detail-item{margin-bottom:10px}.detail-label{font-weight:600;color:#183460}.detail-value{color:#334155}.close-note{font-size:13px;color:#64748b;margin-top:20px;font-style:italic}</style></head><body><div class="container"><div class="icon">✓</div><h1 class="title">Approved Successfully</h1><div class="details"><div class="detail-item"><span class="detail-label">Project:</span><span class="detail-value">' + project + '</span></div><div class="detail-item"><span class="detail-label">Layer:</span><span class="detail-value">' + layerDisplay + '</span></div><div class="detail-item"><span class="detail-label">Date:</span><span class="detail-value">' + getGMT7Time() + '</span></div></div><p class="close-note">You can safely close this page.</p></div><script>setTimeout(function(){window.close();},5000);</script></body></html>';
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createRejectionSuccessPage(name, project, layer) {
  var layerNames = { "FIRST_LAYER": "First Layer", "SECOND_LAYER": "Second Layer", "THIRD_LAYER": "Third Layer" };
  var layerDisplay = layerNames[layer] || layer;
  
  var html = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#dc2626 0%,#b91c1c 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%}.icon{font-size:64px;margin-bottom:20px;color:#dc2626}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#dc2626}.details{background:#f8fafc;padding:20px;border-radius:10px;margin:20px 0;text-align:left;border-left:4px solid #dc2626}.detail-item{margin-bottom:10px}.detail-label{font-weight:600;color:#dc2626}.detail-value{color:#334155}.close-note{font-size:13px;color:#64748b;margin-top:20px;font-style:italic}</style></head><body><div class="container"><div class="icon">✓</div><h1 class="title">Rejection Submitted</h1><div class="details"><div class="detail-item"><span class="detail-label">Project:</span><span class="detail-value">' + project + '</span></div><div class="detail-item"><span class="detail-label">Layer:</span><span class="detail-value">' + layerDisplay + '</span></div><div class="detail-item"><span class="detail-label">Date:</span><span class="detail-value">' + getGMT7Time() + '</span></div></div><p class="close-note">Notifications sent. You can safely close this page.</p></div><script>setTimeout(function(){window.close();},5000);</script></body></html>';
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createErrorPage(message) {
  var html = '<!DOCTYPE html><html><head><style>@import url("https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap");body{font-family:"Inter",sans-serif;text-align:center;padding:20px;background:linear-gradient(135deg,#326BC6 0%,#183460 100%);color:white;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.container{background:white;color:#333;padding:40px 30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.15);max-width:500px;width:90%}.icon{font-size:64px;margin-bottom:20px;color:#dc2626}.title{font-weight:700;font-size:28px;margin-bottom:8px;color:#dc2626}.message{background:linear-gradient(135deg,#fef2f2 0%,#fee2e2 100%);padding:18px;border-radius:8px;margin:20px 0;border-left:4px solid #dc2626;text-align:left}</style></head><body><div class="container"><div class="icon">✕</div><h1 class="title">Error</h1><div class="message">' + message + '</div><p>Please try again or contact support.</p></div></body></html>';
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// HELPER - Handle rejection from web form
// ============================================

function handleRejectionFromWeb(params) {
  return updateRejectionStatus(params.name, params.email, params.project, params.layer, params.rejectionNote);
}

// ============================================
// MENU & UTILITIES
// ============================================

function showSummary(results, count) {
  if (count > 0) {
    var message = "Successfully sent " + count + " approval(s):\n\n";
    results.forEach(function(r) {
      message += "• " + r.project + " → " + r.layer + " (" + r.email + ")\n";
    });
    SpreadsheetApp.getUi().alert("Approvals Sent", message);
  } else {
    SpreadsheetApp.getUi().alert("No Action", "No approvals to send at this time");
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Multi-Layer Approval')
    .addItem('Send Approvals', 'sendMultiLayerApproval')
    .addItem('Manual Trigger Next Layer', 'autoTriggerNextLayer')
    .addSeparator()
    .addItem('View Logs', 'viewLogs')
    .addToUi();
}

function viewLogs() {
  SpreadsheetApp.getUi().alert("Logs", "Go to Extensions > Apps Script > Executions to view logs");
}