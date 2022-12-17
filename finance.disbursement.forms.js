function onSubmit(event) {
  const responseId = event.response.getId();
  Logger.log("responseId=%s", responseId);

  var date = new Date();

  var day = Utilities.formatDate(date, "GMT+1", "yyyy/MM");
  var docfolder = Utilities.formatString("/Auslagen/%s/", day);
  var folder = mkFolder(DriveApp.getRootFolder(), docfolder);

  var time = Utilities.formatDate(date, "GMT+1", "yyyyMMdd-HHmmss");
  var fin = Utilities.formatString("FIN-%s", time);

  var doc = createDocument(responseId, folder, fin);
  var docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);

  // sendNotification(docFile);
}

/**
 * @param {DocumentApp.Document} doc
 */
function sendNotification(doc) {
  MailApp.sendEmail({
    to: "finace@fuks.org",
    subject: "Auslage " + doc.getName(),
    htmlBody: "" +
      "Ein eine Auslage wurde beantragt. " +
      "Zu finden ist diese unter <a href=\"" + doc.getUrl() + "\">" + doc.getName() + "</a>."
  });
}

/**
 * @param {string} responseId
 * @param {DriveApp.Folder} folder
 * @param {string} filename
 */
function createDocument(responseId, folder, filename) {
  const active = FormApp.getActiveForm();
  const response = active.getResponse(responseId);
  const itemResponses = response.getItemResponses();

  const doc = DocumentApp.create(filename);
  const body = doc.getBody();

  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#666666";

  const header = doc.addHeader();
  header.appendParagraph(response.getRespondentEmail()).setAttributes(style);

  const title = body.appendParagraph(filename);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);

  for (var inx = 0; inx < itemResponses.length; inx++) {
    const itemResponse = itemResponses[inx];
    const item = itemResponse.getItem();
    const title = item.getTitle();
    const type = item.getType();

    Logger.log("title=%s type=%s", title, type);

    if (type == FormApp.ItemType.TEXT || type == FormApp.ItemType.PARAGRAPH_TEXT) {
      const headline = body.appendParagraph(title);
      headline.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(itemResponse.getResponse())
    }

    if (type == FormApp.ItemType.IMAGE) {
      const image = item.asImageItem();
      body.appendImage(image);
    }

    if (type == FormApp.ItemType.FILE_UPLOAD) {
      const headline = body.appendParagraph(title);
      headline.setHeading(DocumentApp.ParagraphHeading.HEADING2);

      const fileIds = itemResponse.getResponse();
      const attachments = mkFolder(folder, filename + " (Attachments)");

      for (const fileId of fileIds) {
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        Logger.log("fileId=%s mimeType=%s", fileId, mimeType);

        const paragraph = body.appendParagraph(file.getName());
        paragraph.setLinkUrl(file.getUrl());

        attachments.addFile(file);

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

    return doc;
}

/**
 * @param {DriveApp.Folder} parent
 * @param {string} path
 */
function mkFolder(parent, path) {
    var folder = parent;

    const parts = path.split("/");
    for (const foldername of parts) {
        if (foldername == "") {
            continue;
        }

        var folders = folder.getFoldersByName(foldername);

        if (folders.hasNext()) {
            Logger.log("Folder %s exist", foldername);
            folder = folders.next();
        } else {
            Logger.log("Creating %s", foldername);
            folder = folder.createFolder(foldername);
        }
    }

    return folder;
}