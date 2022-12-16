function onSubmit(event) {
  var responseId = event.response.getId();
  Logger.log("responseId=%s", responseId);

  var active = FormApp.getActiveForm();
  var response = active.getResponse(responseId);
  var itemResponses = response.getItemResponses();

  for (var inx = 0; inx < itemResponses.length; inx++) {
    var itemResponse = itemResponses[inx];
    Logger.log('Response to the question "%s" was "%s"',
               itemResponse.getItem().getTitle(),
               itemResponse.getResponse());
  }
}

function test() {
    var date = new Date();

    var day = Utilities.formatDate(date, "GMT+1", "yyyy-MM-dd");
    var docfolder = Utilities.formatString("/finance/%s/", day);
    var folder = mkdirFolders(docfolder);

    var time = Utilities.formatDate(date, "GMT+1", "yyyy-MM-dd.HH-mm-ss");
    var fin = Utilities.formatString("FIN-%s", time);

    var doc = createDocument(fin);
    var docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
}

function createDocument(fin) {
    var doc = DocumentApp.create(fin);
    var body = doc.getBody();

    // Append a paragraph, with heading 1.
    var title = body.appendParagraph(fin);
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);

    return doc;
}

function mkdirFolders(path) {
    var parts = path.split("/");
    var folder = DriveApp.getRootFolder();

    for (const part of parts) {
        if (part == "") {
            continue;
        }

        folder = mkdirFolder(folder, part);
    }

    return folder;
}

function mkdirFolder(parent, foldername) {
    var folders = parent.getFoldersByName(foldername);
    var folder;

    if (folders.hasNext()) {
        Logger.log("Folder %s exist", foldername);
        folder = folders.next();
    } else {
        Logger.log("Creating %s", foldername);
        folder = parent.createFolder(foldername);
    }

    return folder;
}
