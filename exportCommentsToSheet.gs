function processDocsByFileId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("File IDs");
  var fileIds = sheet.getRange("A1:A").getValues().flat().filter(Boolean); // Filter out undefined values
  var timestamp = new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '');
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(timestamp);
  Logger.log("New sheet created with name: " + timestamp);
  fileIds.forEach(fileId => {
    if (fileId) {
      Logger.log("Processing file ID: " + fileId);
      processDocument(fileId, newSheet);
    }
  });
}

function processDocument(fileId, newSheet) {
  try {
    var doc = DocumentApp.openById(fileId);
    var body = doc.getBody();
    var text = body.getText();
    var wordCount = text.trim().split(/\s+/).length;
    Logger.log("File ID: " + fileId + ", Word Count: " + wordCount);
    exportCommentsToSheet(fileId, newSheet);
  } catch (error) {
    Logger.log("Error processing file ID " + fileId + ": " + error);
  }
}

function exportCommentsToSheet(fileId, sheet) {
  var doc = DocumentApp.openById(fileId);
  var file = Drive.Files.get(fileId, {fields: 'owners'});
  var owners = file.owners;
  var ownerEmail = owners[0].emailAddress;
  var ownerName = owners[0].displayName;
  var docUrl = doc.getUrl();
  var docTitle = doc.getName();
  Logger.log("Exporting comments for document: " + docTitle);
  var comments = retrieveComments(fileId);
  var data = [['Document URL', 'Document Title', 'Owner Name', 'Reviewer', 'Text', 'Status', 'createdDate']];
  for (var i = 0; i < comments.length; i++) {
    var comment = comments[i];
    var reviewer = comment.author; // Changed 'author' to 'reviewer'
    var text = comment.content;
    var status = comment.status;
    var timestamp = comment.createdDate;
    data.push([docUrl, docTitle, ownerName, reviewer, text, status, timestamp]);
  }
  if (sheet.getLastRow() === 0) { // Check if the sheet is empty
    sheet.appendRow(data[0]); // Append header only if the sheet is empty
  }
  for (var i = 1; i < data.length; i++) { // Append data rows individually
    sheet.appendRow(data[i]);
  }
}

function retrieveComments(fileId) {
  var info = [];
  var callArguments = {'maxResults': 100, 'fields': 'items(commentId,content, status, author, createdDate),nextPageToken'};
  var docComments, pageToken;
  do {
    callArguments['pageToken'] = pageToken;
    docComments = Drive.Comments.list(fileId,callArguments);
    info = info.concat(getCommentsInfo(docComments.items));
    pageToken = docComments.nextPageToken;
    Logger.log("Retrieved " + info.length + " comments so far.");
  } while(pageToken);
  return(info);
}

function getCommentsInfo(items) {
  var commentInfo = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    commentInfo.push({
      commentId: item.commentId,
      content: item.content,
      status: item.status,
      author: item.author.displayName, // Changed 'author' to 'reviewer' here if needed in other places
      createdDate: item.createdDate,
    });
  }
  return commentInfo;
}