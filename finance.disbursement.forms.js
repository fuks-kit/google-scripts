const auslagenFolderId = "13uwc4LBe3DNMEJvCBLG2nRQBHv8tkyUM"

function onSubmit(event) {
  const responseId = event.response.getId();
  Logger.log("responseId=%s", responseId);

  let date = new Date();

  let time = Utilities.formatDate(date, "GMT+1", "yyyyMMdd-HHmmss");
  let fin = Utilities.formatString("FIN-%s", time);

  const active = FormApp.getActiveForm();
  const response = active.getResponse(responseId);

  const name = response.getRespondentEmail().replace("@fuks.org", "");

  let day = Utilities.formatDate(date, "GMT+1", "yyyy/MM");
  let docfolder = Utilities.formatString("/%s/%s_%s", day, fin, name);
  let folder = mkFolder(DriveApp.getFolderById(auslagenFolderId), docfolder);

  let doc = createDocument(response, folder, fin);
  let docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);

  MailApp.sendEmail({
    to: "vorstand-finanzen-recht@fuks.org",
    cc: response.getRespondentEmail(),
    subject: Utilities.formatString("Auslage %s (%s)", name, docFile.getName()),
    htmlBody: "" +
      "Ein eine Auslage wurde beantragt. " +
      "Zu finden ist diese unter <a href=\"" + doc.getUrl() + "\">" + docFile.getName() + "</a>."
  });
}

/**
 * @param {FormApp.FormResponse} response
 * @param {DriveApp.Folder} folder
 * @param {string} filename
 */
function createDocument(response, folder, filename) {
  const itemResponses = response.getItemResponses();

  const doc = DocumentApp.create(filename);
  const body = doc.getBody();

  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#666666";

  const header = doc.addHeader();
  header.appendParagraph("Antragsteller: " + response.getRespondentEmail()).setAttributes(style);

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
      body.appendParagraph(itemResponse.getResponse());
      body.appendParagraph("\n");
    }

    if (type == FormApp.ItemType.IMAGE) {
      const image = item.asImageItem();
      body.appendImage(image);
      body.appendParagraph("\n");
    }

    if (type == FormApp.ItemType.FILE_UPLOAD) {
      const headline = body.appendParagraph(title);
      headline.setHeading(DocumentApp.ParagraphHeading.HEADING2);

      const fileIds = itemResponse.getResponse();
      for (const fileId of fileIds) {
        //
        // Moving an attachment is not possible currently with Drive.
        // You need to create a copy of the attachment.
        //
        const original = DriveApp.getFileById(fileId);
        const copy = folder.createFile(original.getBlob());

        const mimeType = copy.getMimeType();
        Logger.log("fileId=%s mimeType=%s", fileId, mimeType);

        const paragraph = body.appendParagraph(copy.getName());
        paragraph.setLinkUrl(copy.getUrl());

        // Remove attachments from forms
        /*
        let parents = original.getParents();
        while (parents.hasNext()) {
          let parent = parents.next();
          Logger.log("parent=%s --> %s", parent.getId(), parent.getName());
          parent.removeFile(original);
        }
        */

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

      body.appendParagraph("\n");
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