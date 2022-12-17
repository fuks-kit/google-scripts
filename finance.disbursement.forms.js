function onSubmit(event) {
  const responseId = event.response.getId();
  Logger.log("responseId=%s", responseId);

  const active = FormApp.getActiveForm();
  const response = active.getResponse(responseId);
  const itemResponses = response.getItemResponses();

  var doc = DocumentApp.create("zzzzzzzzz-testfile");
  var body = doc.getBody();

  for (var inx = 0; inx < itemResponses.length; inx++) {
    const itemResponse = itemResponses[inx];
    const item = itemResponse.getItem();
    const title = item.getTitle();
    const type = item.getType();

    Logger.log("title=%s type=%s response=%s", title, type);

    if (type == FormApp.ItemType.TEXT || type == FormApp.ItemType.PARAGRAPH_TEXT) {
      var headline = body.appendParagraph(title);
      headline.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(itemResponse.getResponse())
    }

    if (type == FormApp.ItemType.IMAGE) {
      const image = item.asImageItem();
      body.appendImage(image);
    }

    if (type == FormApp.ItemType.FILE_UPLOAD) {
      var headline = body.appendParagraph(title);
      headline.setHeading(DocumentApp.ParagraphHeading.HEADING2);

      const fileIds = itemResponse.getResponse();

      for (const fileId of fileIds) {
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        Logger.log("fileId=%s mimeType=%s", fileId, mimeType);

        const paragraph = body.appendParagraph(file.getName());
        paragraph.setLinkUrl(file.getUrl());

        // if (mimeType == "image/jpeg") {
        //   const img = file.getBlob();
        //   body.appendImage(img);
        // } else {
        //   const paragraph = body.appendParagraph(file.getName());
        //   paragraph.setLinkUrl(file.getUrl());
        // }

        // Converting from image/heif to image/jpeg is not supported.
        // if (mimeType.startsWith("image")) {
        //   const jpeg = file.getAs("image/jpeg");
        //   body.appendImage(jpeg);
        // }
      }
    }
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
