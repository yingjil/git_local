function processDocsByFileId() {
  // Get the file IDs from a spreadsheet (adjust the sheet and range as needed)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("File IDs"); // Replace with your sheet name
  var fileIds = sheet.getRange("A1:A").getValues().flat(); 

  fileIds.forEach(fileId => {
    if (fileId) { // Skip empty cells
      processDocument(fileId);
    }
  });
}

function processDocument(fileId) {
  try {
    // Get the document
    var doc = DocumentApp.openById(fileId); 

    // Perform actions on the document
    // Example: Count and log the number of words
    var body = doc.getBody();
    var text = body.getText();
    var wordCount = text.trim().split(/\s+/).length;
    Logger.log("File ID: " + fileId + ", Word Count: " + wordCount);

    exportCommentsToSheet(fileId);

  } catch (error) {
    Logger.log("Error processing file ID " + fileId + ": " + error);
  }
}

function exportCommentsToSheet(fileId) {
  var doc = DocumentApp.openById(fileId);
  var file = Drive.Files.get(fileId, {fields: 'owners'}); 
  var owners = file.owners;
  var ownerEmail = owners[0].emailAddress;
  var ownerName = owners[0].displayName;

  var docUrl = doc.getUrl();
  var docTitle = doc.getName(); 
  var comments = retrieveComments(fileId);
  var sheet = SpreadsheetApp.create('Comments Export');
  var data = [['Document URL', 'Document Title', 'Owner Name', 'Author', 'Text', 'Status', 'createdDate']];

  for (var i = 0; i < comments.length; i++) {
    var comment = comments[i];
    var author = comment.author;
    var text = comment.content;
    var status = comment.status;
    var timestamp = comment.createdDate;
   
    data.push([docUrl, docTitle, ownerName, author, text, status, timestamp]);
  }
  var timestamp = new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(timestamp);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
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
      author: item.author.displayName,
      createdDate: item.createdDate,
    });
  }
  return commentInfo;
}