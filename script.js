function converterGoogleDocParaHtmlLimpo() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var output = [];
  var images = [];
  var listCounters = {};

  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    output.push(processItem(child, listCounters, images));
  }

  var html = output.join('\r');
  var htmlFileId = saveHtmlToDrive(html); // Salve o arquivo HTML no Google Drive
  enviarEmailComLink(htmlFileId, images);
}

function saveHtmlToDrive(html) {
  var blob = Utilities.newBlob(html, 'text/html', 'document.html');
  var file = DriveApp.createFile(blob);
  return file.getId();
}

function enviarEmailComLink(htmlFileId, imagens) {
  var htmlLink = 'Link para o arquivo HTML no Google Drive: ' + getGoogleDriveLink(htmlFileId);
  
  // Anexar apenas as imagens ao e-mail
  var anexos = [];
  for (var j = 0; j < imagens.length; j++) {
    anexos.push({
      "fileName": imagens[j].nome,
      "mimeType": imagens[j].tipo,
      "content": imagens[j].blob.getBytes()
    });
  }

  var assunto = DocumentApp.getActiveDocument().getName() + " - HTML e Imagens";
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: assunto,
    body: htmlLink, // Inclua o link para o arquivo HTML no corpo do e-mail
    inlineImages: {},
    attachments: anexos
  });
}

function getGoogleDriveLink(fileId) {
  return 'https://drive.google.com/uc?export=download&id=' + fileId;
}

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "";
  var suffix = "";

  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
      case DocumentApp.ParagraphHeading.HEADING6:
        prefix = "<h6>";
        suffix = "</h6>";
        break;
      case DocumentApp.ParagraphHeading.HEADING5:
        prefix = "<h5>";
        suffix = "</h5>";
        break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>";
        suffix = "</h4>";
        break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>";
        suffix = "</h3>";
        break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>";
        suffix = "</h2>";
        break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>";
        suffix = "</h1>";
        break;
      default:
        prefix = "<p>";
        suffix = "</p>";
    }

    if (item.getNumChildren() == 0) return "";
  } else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
    processImage(item, images, output);
  } else if (item.getType() === DocumentApp.ElementType.LIST_ITEM) {
    var listItem = item;
    var gt = listItem.getGlyphType();
    var key = listItem.getListId() + '.' + listItem.getNestingLevel();
    var counter = listCounters[key] || 0;

    // Rest of the code here...
  }

  output.push(prefix);

  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output);
  } else {
    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();

      // Walk through all the child elements of the doc.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child, listCounters, images));
      }
    }
  }

  output.push(suffix);
  return output.join('');
}

function processText(item, output) {
  var text = item.getText();
  var indices = item.getTextAttributeIndices();

  if (indices.length <= 1) {
    // Assuming that a whole para fully italic is a quote
    if (item.isBold()) {
      output.push('<b>' + text + '</b>');
    } else if (item.isItalic()) {
      output.push('<blockquote>' + text + '</blockquote>');
    } else if (text.trim().indexOf('http://') == 0) {
      output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
    } else {
      output.push(text);
    }
  } else {
    for (var i = 0; i < indices.length; i++) {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i + 1 < indices.length ? indices[i + 1] : text.length;
      var partText = text.substring(startPos, endPos);

      Logger.log(partText);

      if (partAtts.ITALIC) {
        output.push('<i>');
      }
      if (partAtts.BOLD) {
        output.push('<b>');
      }
      if (partAtts.UNDERLINE) {
        output.push('<u>');
      }

      // If someone has written [xxx] and made this whole text some special font, like superscript
      // then treat it as a reference and make it superscript.
      // Unfortunately in Google Docs, there's no way to detect superscript
      if (partText.indexOf('[') == 0 && partText[partText.length - 1] == ']') {
        output.push('<sup>' + partText + '</sup>');
      } else if (partText.trim().indexOf('http://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      } else {
        output.push(partText);
      }

      if (partAtts.ITALIC) {
        output.push('</i>');
      }
      if (partAtts.BOLD) {
        output.push('</b>');
      }
      if (partAtts.UNDERLINE) {
        output.push('</u>');
      }
    }
  }
}

function processImage(item, images, output) {
  images = images || [];
  var blob = item.getBlob();
  var contentType = blob.getContentType();
  var extension = "";
  if (/\/png$/.test(contentType)) {
    extension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    extension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    extension = ".jpg";
  } else {
    throw "Unsupported image type: " + contentType;
  }
  var imagePrefix = "Image_";
  var imageCounter = images.length;
  var nome = imagePrefix + imageCounter + extension;
  imageCounter++;
  output.push('<img src="cid:' + nome + '" />');
  images.push({
    "blob": blob,
    "tipo": contentType,
    "nome": nome
  });
}