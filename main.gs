/* This function converts the google document into mark down and saves it in the common folder as configured by the user.
It will also save any inline images in the same folder where converted MD file is placed */
function ConvertToMarkdown(doc, index_start, index_end) {
  // initialize function parameters
  if (doc === undefined) doc = DocumentApp.getActiveDocument();
  if (index_start === undefined) index_start = 0;
  var body = doc.getBody();
  var numChildren = body.getNumChildren();
  if (index_end === undefined) index_end = numChildren;
  
  var text = "";
  var inSrc = false;
  var inClass = false;
  var globalImageCounter = 0;
  var globalListCounters = {};
  // edbacher: added a variable for indent in src <pre> block. Let style sheet do margin.
  var srcIndent = "";
  
  var folder = getParentFolder(doc);
  // remove previous MD files
  var filename = doc.getName() + '.md';
  removeExistingFiles(folder, filename);

  var attachments = [];
  var file;
  var commonFolder;
  var folderName;
  var blob;
  var content_type;
  var suffix;
  var image_name;
  var photo;
  
//  try {  
  var prev_type = 'PARAGRAPH';
  var prev_heading = 'Normal';
  var child_heading = '';
//  var isCode = false;
  
  // Walk through all the child elements of the doc.
  for (var i = index_start; i < index_end; i++) {
    var child = body.getChild(i);
    
    // check the child position
    var child_type = child.getType();
    if (child_type == 'UNSUPPORTED') continue;
    
    if (child_type == 'PARAGRAPH') {
      child_heading = child.getHeading();
    } else {
      child_heading = '';
    }
    //      Logger.log('type: ' + child_type + '/' + child_heading);
    
    var result = processParagraph(i, child, inSrc, globalImageCounter, globalListCounters);
    
    globalImageCounter += (result && result.images) ? result.images.length : 0;
    
    
    if (result !== null) {
      inSrc = result.inSrc;
//      Logger.log('inSrc: ' + inSrc);
      
      if (inClass) {
        text += result.text + "\n\n";
      } else if (inSrc) {
//        text += (srcIndent + escapeHTML(result.text) + "\n");
        text += result.text + '\n';
        
      } else if (result.text && result.text.length > 0) {
//        Logger.log(child_type + '/' + child_heading + '**' + prev_type + '/' + prev_heading);
        if (child_type == 'PARAGRAPH') {
          if (prev_type == 'LIST_ITEM') {
            text += '\n';
          }
          if (child_heading == 'Normal' && result.text.charAt(0) != '#') {
            text += result.text + '\n\n';
          } else {
            text += result.text + '\n';
          }
        } else if (child_type == 'LIST_ITEM') {
          text += result.text + '\n';
        } else {
          text += result.text + '\n\n';
        }
        // set previous type info
        prev_type = child_type;
        prev_heading = child_heading;        
      
      }
      
      if (result.images && result.images.length > 0) {
        for (var j = 0; j < result.images.length; j++) {
          attachments.push({
            "fileName": result.images[j].name,
            "mimeType": result.images[j].type,
            "content": result.images[j].bytes
          });
        }
      }
      
    } else if (inSrc) { // support empty lines inside source code
      text += '\n';
    }
  }
  
  file = DriveApp.createFile(filename, text, 'text/plain');
  folder.addFile(file)
  
  //If there are any attachments in the file, it has to be saved in the same directory.
  //Due to this issue [http://code.google.com/p/google-apps-script-issues/issues/detail?id=1239], image files are created using blob. and replaced using Drive API.
  if (attachments.length > 0) {
    for (var iterator = 0; iterator < attachments.length; iterator++) {
      blob = attachments[iterator].content;
      content_type = blob.getContentType()
      Logger.log('content_type: ' + content_type);
      suffix = content_type.split("/")[1] //e.g gif/jpg or png
      blob.setName('test.png');  // Invent a name, the blob seems to need it?
      
//      try {
      var image_name = "image_" + iterator + "." + suffix;
      Logger.log('image_name: ' + image_name);
      removeExistingFiles(folder, image_name);
      photo = folder.createFile(blob);
      photo.setName(image_name);
//      } catch (e) {
//        throw ("Error in saving attached images : " + e);
//      }
      
    }
  }
//  } catch (e) {
//    var errorMsg = "";
//    //While displaying error message, we display the last converted text, so that the users can know, after which line the conversion failed.
//    //Check if there is any last converted text. If so take the last sentence from the converted text. If not, just display the error message.
//    if (text != null && text.length != 0 && text.trim() !== "") {
//      var sentence = text.split(".");
//      if (sentence.length > 1) {
//        errorMsg = "Error after the line : \"" + sentence[sentence.length - 2] + "\".\n\n" + e;
//      } else if (sentence.length == 1) {
//        errorMsg = "Error after the line : \"" + sentence[sentence.length - 1] + "\".\n\n" + e;
//      } else if (sentence.length == 0) {
//        errorMsg = "Error after the text : \"" + text + "\".\n\n" + e;;
//      }
//    } else {
//      errorMsg = e;
//    }
//    //Showing the error message in alert window.
//    DocumentApp.getUi().alert("Error", errorMsg, DocumentApp.getUi().ButtonSet.OK);
//  }
//  return file;
}


/* This function converts the current open document into an MD file and saves it in the common folder. It will then show you the pop up to download the MD file */
function downloadMdFile() {
  var file = ConvertToMarkdown();
  
  DocumentApp.getUi().showDialog(
    HtmlService
    .createHtmlOutput('<a href="https://docs.google.com/a/fusioncharts.com/uc?export=download&id=' + file.getId() + '">Download the converted MD file</a>')
  .setTitle('Download Link for MD file')
  .setWidth(400 /* pixels */ )
  .setHeight(100 /* pixels */ ));
}

/* This function returns the current folder. */
function getParentFolder(doc) {
  // get the parent folders, could be more than one
  var directParents = DriveApp.getFileById(doc.getId()).getParents();
  
  var folderCount = 0;
  while(directParents.hasNext()) {
    folderCount++;
    var currentFolder = directParents.next();
  }
  if (folderCount == 1) {
    return currentFolder;
  }
  
  Logger.log('Found more than one parent folder.');
  return null;
}

/* This function removes previous verions of the file */
function removeExistingFiles(folder, filename) {
  Logger.log(folder.getName());
  Logger.log(filename)
  var files = folder.getFilesByName(filename);
  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
}

function escapeHTML(text) {
  return text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Process each child element (not just paragraphs).
function processParagraph(index, element, inSrc, imageCounter, listCounters) {
//  Logger.log('element type: ' + element.getType());
  // First, check for things that require no processing.
  if (element.getNumChildren() == 0) {
    return null;
  }
  // Punt on TOC.
  if (element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
    return {
      "text": "[[TOC]]"
    };
  }
  
  // Set up for real results.
  var result = {};
  result.inSrc = inSrc;
  var pOut = "";
  var textElements = [];
  var imagePrefix = "image_";
  var child_heading = '';
  var cell_text = '';
  
  // Start source code block
  if (!inSrc) {
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      if (element.getText().slice(0, 3) == '```') {
        inSrc = true;
      }
    }
  }
  
  // Process Table
  if (element.getType() === DocumentApp.ElementType.TABLE) {
    var nRows = element.getNumChildren();
    var nCols = element.getChild(0).getNumCells();
    for (var row = 0; row < nRows; row++) {
      textElements.push("|");
      // process this row
      for (var col = 0; col < nCols; col++) {
        var cell = element.getChild(row).getChild(col);
        
        var num_para = cell.getNumChildren();
        
        // add initial space
        textElements.push(' ');
        
        // process each paragraph
        for (var para_num = 0; para_num < num_para; para_num++) {          
          var cell_elem = cell.getChild(para_num);
          var cell_elem_type = cell_elem.getType();
          
          // add break between paragraphs
          if (para_num > 0) textElements.push('<br>');
        
          if (cell_elem_type == 'PARAGRAPH') {
            var cell_result = processParagraph(index, cell_elem, inSrc, imageCounter, listCounters);           
            
            if (cell_result) {
              cell_text = cell_result.text;
              
              // remove bold from heading paragraphs
              if (row == 0) {
                cell_text = cell_text.replace(/\*\*/g, '');
              }
            } else {
              cell_text = ' ';
            }
            textElements.push(cell_text);
          } else {
            Logger.log('unsupported element: ' + cell_elem_type);
            textElements.push('<NOT_PARA>');
          }
        }
        
        // add final space and cell divider
        textElements.push(' |');
      }
      
      // add new line, except for last row
      if (row < nRows - 1) {
        textElements.push("\n");
      }
      
      // process the delimiter row
      if (row == 0) {
        textElements.push("|");
        // process this row
        for (var col = 0; col < nCols; col++) {
          textElements.push(" :--- |");
        }
        textElements.push("\n");
      }
    }

  }
  
  // Process various types (ElementType).
  for (var i = 0; i < element.getNumChildren(); i++) {
    var t = element.getChild(i).getType();
//    Logger.log('element child type: ' + t);
    if (t === DocumentApp.ElementType.TABLE_ROW) {
      // do nothing: already handled TABLE_ROW
    } else if (t === DocumentApp.ElementType.TEXT) {
      var txt = element.getChild(i);
      pOut += txt.getText();
      textElements.push(txt);
    } else if (t === DocumentApp.ElementType.INLINE_IMAGE) {
      result.images = result.images || [];
      var img = element.getChild(i);
      var blob = img.getBlob();
      var alt_title = img.getAltTitle() || 'no alt title';
      
      Logger.log('alt_title: ' + alt_title);
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
      var name = imagePrefix + imageCounter + extension;
      imageCounter++;
      textElements.push('![' + alt_title + '](' + name + ')');
      result.images.push({
        "bytes": blob,
        "type": contentType,
        "name": name
      });
    } else if (t === DocumentApp.ElementType.PAGE_BREAK) {
      // ignore
    } else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {
//      textElements.push('* * *\n');
      textElements.push('-------');
    } else if (t === DocumentApp.ElementType.FOOTNOTE) {
      textElements.push(' (NOTE: ' + element.getChild(i).getFootnoteContents().getText() + ')');
    } else {
      //throw "Paragraph "+index+" of type "+element.getType()+" has an unsupported child: "
      //+t+" "+(element.getChild(i)["getText"] ? element.getChild(i).getText():'')+" index="+result;
      throw "Unsupported format in current file :" + t + " " + (element.getChild(i)["getText"] ? element.getChild(i).getText() : '') + ". Cannot be converted into an MD file";
    }
  }
  
  if (textElements.length == 0) {
    // Isn't result empty now?
    return result;
  }
    
  prefix = findPrefix(inSrc, element, listCounters);
  
  var pOut = '';
  for (var i = 0; i < textElements.length; i++) {
    pOut += processTextElement(inSrc, textElements[i]);
  }
  
  // correct multiple, concatenated bold formatting
  pOut = pOut.replace(/\*\*\*\*/g, '');
  
  // remove bold formatting from headings
  if (prefix.charAt(0) == '#' || pOut.slice(0, 3) == '**#') {
    if (pOut.slice(0, 2) == '**' && pOut.slice(-2) == '**') {
      pOut = pOut.slice(2, -2);
    }
  }
  
  // replace Unicode quotation marks
  pOut = pOut.replace('\u201d', '"').replace('\u201c', '"');
  
  // replace smart quotes here
  
  // replace non-breaking spaces
  pOut = pOut.replace('\u00a0', ' ');
  
  // remove heading numbers on appendix headings
  if (prefix.charAt(0) == '#') {
    var heading = pOut;
    var matches = /^[\d\.]+\s(Appendix.*)/.exec(pOut);
    if (matches) {
      heading = matches[1];
    }
    pOut = heading;
  }

  // End source code block
  if (result.inSrc) {  // check the original so that you don't immediately turn off the block
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      if (element.getText().slice(0, 3) == '```') {
        inSrc = false;
      }
    }
  }
  
  // setup return variable
  result.text = prefix + pOut;
  result.inSrc = inSrc;
  
  // return result
  return result;
}

// Add correct prefix to list items.
function findPrefix(inSrc, element, listCounters) {
  var prefix = "";
  if (!inSrc) {
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var paragraphObj = element;
      
      switch (paragraphObj.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
        case DocumentApp.ParagraphHeading.HEADING6:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING5:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING4:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING3:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING2:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING1:
          prefix += "# ";
        default:
      }
    } else if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
      var listItem = element;
      var nesting = listItem.getNestingLevel()
      for (var i = 0; i < nesting; i++) {
        prefix += "    ";
      }
      var gt = listItem.getGlyphType();
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix += "* ";
      } else {
        // Ordered list (<ol>):
        var key = listItem.getListId() + '.' + listItem.getNestingLevel();
        var counter = listCounters[key] || 0;
        counter++;
        listCounters[key] = counter;
        prefix += counter + ". ";
      }
    }
  }
  return prefix;
}

function processTextElement(inSrc, txt) {
  if (typeof(txt) === 'string') {
    return txt;
  }
  
  var pOut = txt.getText();
  if (!txt.getTextAttributeIndices) {
    return pOut;
  }
  
  var attrs = txt.getTextAttributeIndices();
  var lastOff = pOut.length;
  
  for (var i = attrs.length - 1; i >= 0; i--) {
    var off = attrs[i];
    var url = txt.getLinkUrl(off);
    var font = txt.getFontFamily(off);
    
    if (url) { // start of link
      if (i >= 1 && attrs[i - 1] == off - 1 && txt.getLinkUrl(attrs[i - 1]) === url) {
        // detect links that are in multiple pieces because of errors on formatting:
        i -= 1;
        off = attrs[i];
        url = txt.getLinkUrl(off);
      }
      pOut = pOut.substring(0, off) + '[' + pOut.substring(off, lastOff) + '](' + url + ')' + pOut.substring(lastOff);
    } else if (font) {
      if (!inSrc && (font == 'Source Code Pro' || font == 'Consolas')) {
        var code_font = font;
        while (i >= 1 && txt.getFontFamily(attrs[i - 1]) && txt.getFontFamily(attrs[i - 1]) == code_font) {
          // detect fonts that are in multiple pieces because of errors on formatting:
          i -= 1;
          off = attrs[i];
        }
        pOut = pOut.substring(0, off) + '`' + pOut.substring(off, lastOff) + '`' + pOut.substring(lastOff);
      }
    }
    
//    Logger.log(txt.getText() + ' //' + off + ': ' + txt.isBold() + '/' + txt.isBold(off));
    if (txt.isBold(off)) {
      var d1 = d2 = "**";
      if (txt.isItalic(off)) {
        // edbacher: changed this to handle bold italic properly.
        d1 = "**_";
        d2 = "_**";
      }
      pOut = pOut.substring(0, off) + d1 + pOut.substring(off, lastOff) + d2 + pOut.substring(lastOff);
    } else if (txt.isItalic(off)) {
      pOut = pOut.substring(0, off) + '_' + pOut.substring(off, lastOff) + '_' + pOut.substring(lastOff);
    }
    
    lastOff = off;
  }

  // remove double toggle italics
  pOut = pOut.replace(/__/g, '');
  
  return pOut;
}
