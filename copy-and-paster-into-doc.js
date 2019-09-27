/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Open DOC2HTML', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('DOC2HTML');
  DocumentApp.getUi().showSidebar(ui);
}

function getMyEmailAddress() {
  var current_user = Session.getActiveUser().getEmail()
  return current_user;
}

// This function converts Google Doc to HTML and emails it to the recipients. 
function ConvertGoogleDocToHtml(to_recipients, cc_recipients, bcc_recipients) {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  
  // Output is where all the HTML  elements get pused onto.
  var output = [];
  var images = [];
  var listCounters = {};
  var success_msg = "Email has been sent to designated recipient(s) successfully!"

  // The entire newsletter is inside a table. The opening table tag is added first.
  output.push('<table bgcolor="#fff" align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;">')
  
  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    
    // Get the nth child element 
    var child = body.getChild(i);

    // Push each HTML-converted element onto the output variable
    // processItem is what converts paragraph styles from the Google Doc into HTML elements
    output.push(processItem(child, listCounters, images));
  }

  // This is end of the email. This is the closing table tag.
  output.push('</table>');
  
  // ??? What is this joining on?
  var html = output.join('\r');

  // Email the HTML-converted elements
  emailHtml(html, images, to_recipients, cc_recipients, bcc_recipients);
  //createDocumentForHtml(html, images);
  return success_msg;
}


// Function emails HTML-converted elements
function emailHtml(html, images, to_recipients, cc_recipients, bcc_recipients) {
  var attachments = [];

  //??? Not sure why images get put into attachment list
  for (var j=0; j<images.length; j++) {
    attachments.push( {
      "fileName": images[j].name,
      "mimeType": images[j].type,
      "content": images[j].blob.getBytes() } );
  }

  var inlineImages = {};
  //??? Is this captioning each image by assigning image name to image blob?
  for (var j=0; j<images.length; j++) {
    inlineImages[[images[j].name]] = images[j].blob;
  }

  // Gets the name of the document
  var name = DocumentApp.getActiveDocument().getName();
  
  // Sends the email
  MailApp.sendEmail({                                                                                
     to: to_recipients,
     cc: cc_recipients,
     bcc: bcc_recipients,
     subject: name,
     htmlBody: html,
     inlineImages: inlineImages
   });
}

//??? Have no clue that this function is going
function createDocumentForHtml(html, images) {
  var name = DocumentApp.getActiveDocument().getName()+".html";
  var newDoc = DocumentApp.create(name);
  newDoc.getBody().setText(html);
  for(var j=0; j < images.length; j++)
    newDoc.getBody().appendImage(images[j].blob);
  newDoc.saveAndClose();
}

//??? Have no clue what this function does
function dumpAttributes(atts) {
  // Log the paragraph attributes.
  for (var att in atts) {
    Logger.log(att + ":" + atts[att]);
  }
}

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "", suffix = "";

  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    // TITLE == TITLE OF DOC (e.g., Deubbuger Newsletter)
    if (item.getHeading() == DocumentApp.ParagraphHeading.TITLE){
      prefix='<tr><td style="text-align:center; bgcolor:#fff; valigh:top; padding:0;"><h1 style="font-size:28px; color:#727272; margin:20px 20px 0; letter-spacing:0.5px" class="email-heading">',
      suffix = "</h1></td></tr>";
    }

    // Subtitle for each section (e.g., The numbers, Lookinging back, Looking forward)
    else if (item.getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
      prefix='<tr><td style="text-align:left; padding: 20px 40px 0 40px"><h2 style="font-size:20px; color:#000000">',
      suffix="</h2></td></tr>";
    }
    
    // minor title (after subtitle) (e.g., "top 5 visited sites, etc")
    else if (item.getHead() == DocumentApp.ParagraphHeading.)
   // NORMAL == default
   else {
    prefix = '<tr><td style="background-color: #fff; padding: 20px 40px 0 40px; font-size: 16px; line-height: 24px; color: #7e7e7e; text-align: left; font-weight:normal;">';
    suffix = "</td></tr>";
  }
  
  if (item.getNumChildren() == 0)
    return "";
}

else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE)
{
  processImage(item, images, output);
}

else if (item.getType()===DocumentApp.ElementType.LIST_ITEM) {
  var listItem = item;
  var gt = listItem.getGlyphType();
  var key = listItem.getListId() + '.' + listItem.getNestingLevel();
  var counter = listCounters[key] || 0;

  // First list item
  if ( counter == 0 ) {
    // Bullet list (<ul>):
    if (gt === DocumentApp.GlyphType.BULLET
        || gt === DocumentApp.GlyphType.HOLLOW_BULLET
        || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
      prefix = '<tr><td style="background-color: #fff; padding: 0px 40px 0 40px; font-size: 16px; line-height: 24px; color: #7e7e7e; text-align: left; font-weight:normal;"><ul style="color: #7e7e7e; font-size: 16px; line-height: 24px; margin: 5px 0 0 0;"><li>', suffix = "</li>";
      }
    else {
      // Ordered list (<ol>):
      prefix = '<tr><td style="background-color: #fff; padding: 0px 40px 0 40px; font-size: 16px; line-height: 24px; color: #7e7e7e; text-align: left; font-weight:normal;"><ol style="color: #7e7e7e; font-size: 16px; line-height: 24px; margin: 5px 0 0 0;"><li>', suffix = "</li>";
    }
  }
  else {
    prefix = "<li>";
    suffix = "</li>";
  }

  if (item.isAtDocumentEnd() || (item.getNextSibling() && (item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM))) {
    if (gt === DocumentApp.GlyphType.BULLET
        || gt === DocumentApp.GlyphType.HOLLOW_BULLET
        || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
      suffix += "</ul></td></tr>";
    }
    else {
      // Ordered list (<ol>):
      suffix += "</ol></td></tr>";
    }

  }

  counter++;
  listCounters[key] = counter;
}

output.push(prefix);

if (item.getType() == DocumentApp.ElementType.TEXT) {
  processText(item, output);
}
else {


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
  if(item.isBold()) {
    output.push('<strong>' + text + '</strong>');
  }
  else if(item.isItalic()) {
    output.push('<blockquote>' + text + '</blockquote>');
  }
  else if (text.trim().indexOf('http://') == 0) {
    output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
  }
  else {
    output.push(text);
  }
}
else {

  for (var i=0; i < indices.length; i ++) {
    var partAtts = item.getAttributes(indices[i]);
    var startPos = indices[i];
    var endPos = i+1 < indices.length ? indices[i+1]: text.length;
    var partText = text.substring(startPos, endPos);

    Logger.log(partText);

    if (partAtts.ITALIC) {
      output.push('<i>');
    }
    if (partAtts.BOLD) {
      output.push('<strong>');
    }
    if (partAtts.UNDERLINE) {
      output.push('<u>');
    }

    // If someone has written [xxx] and made this whole text some special font, like superscript
    // then treat it as a reference and make it superscript.
    // Unfortunately in Google Docs, there's no way to detect superscript
    if (partText.indexOf('[')==0 && partText[partText.length-1] == ']') {
      output.push('<sup>' + partText + '</sup>');
    }
    else if (partText.trim().indexOf('http://') == 0) {
      output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
    }
    else {
      output.push(partText);
    }

    if (partAtts.ITALIC) {
      output.push('</i>');
    }
    if (partAtts.BOLD) {
      output.push('</strong>');
    }
    if (partAtts.UNDERLINE) {
      output.push('</u>');
    }

  }
}
}


function processImage(item, images, output)
{
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
  throw "Unsupported image type: "+contentType;
}
var imagePrefix = "Image_";
var imageCounter = images.length;
var name = imagePrefix + imageCounter + extension;
imageCounter++;
output.push('<img src="cid:'+name+'" width="100%" height="auto" border="0" style="display: block;"/>');
images.push( {
  "blob": blob,
  "type": contentType,
  "name": name});
}