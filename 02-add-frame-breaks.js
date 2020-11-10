// ***************************************************************************
// add-frame-breaks.js
// ***************************************************************************

/**
 * Logs a string to ~/Desktop/Logs/indesign_log.txt
 * 
 * @param {string} s - the string to log
 */
function boop(s) {

  var now = new Date();
  var output = now.toLocaleString() + ': ' + s;
  var logFolder = Folder(Folder.desktop.fsName + '/Logs');

  if (!logFolder.exists) {
    logFolder.create();
  }

  var logFile = File(logFolder.fsName + '/indesign_log.txt');
  logFile.lineFeed = 'Unix';
  logFile.encoding = 'UTF-8';

  logFile.open('a');
  logFile.writeln(output);
  logFile.close();
}


/**
 * Make a text frame that fills the page margins.
 * 
 * @param {Page} page - the page to make the text frame on
 * @returns {TextFrame} - the created text frame
 */
function makeTextFrame(page) {

  var y1 = page.marginPreferences.top;
  var y2 = page.bounds[2] - page.marginPreferences.bottom;

  // figure out the left and right bounds of the box
  var x1;
  var x2;
  if (page.side == PageSideOptions.LEFT_HAND) {
    x1 = page.marginPreferences.right;
    x2 = page.bounds[3] - page.marginPreferences.left;
  } else if (page.side == PageSideOptions.RIGHT_HAND) {
    x1 = page.marginPreferences.left + page.bounds[1];
    x2 = page.bounds[3] - page.marginPreferences.right;
  } else if (page.side == PageSideOptions.SINGLE_SIDED) {
    x1 = page.marginPreferences.left;
    x2 = page.bounds[3] - page.marginPreferences.right;
  }

  return page.textFrames.add(
    undefined,
    undefined,
    undefined,
    {
      geometricBounds: [
        y1, x1, y2, x2
      ]
    }
  );
}

var doc = app.activeDocument;
var PAR_1_STYLE = doc.paragraphStyles.item('age-population-sex');
var PAR_2_STYLE = doc.paragraphStyles.item('disease');
var PAR_3_STYLE = doc.paragraphStyles.item('names');
var PAR_4_STYLE = doc.paragraphStyles.item('space between');

/**
 * Determine if the paragraph is OK to be the last paragraph of the text frame.
 * 
 * @param {TextFrame} textFrame 
 */
Paragraph.prototype.validLastPar = function(textFrame) {

  if (this.appliedParagraphStyle == PAR_4_STYLE) {
    return true;
  }

  if (
    this.appliedParagraphStyle == PAR_1_STYLE ||
    this.appliedParagraphStyle == PAR_2_STYLE
  ) {
    return false;
  }

  if (this.appliedParagraphStyle == PAR_3_STYLE) {
    var parLastLine = this.lines.lastItem();
    return textFrame.lines.lastItem() == parLastLine;
  }

  return true;

};

// replace all spaces that do not follow semicolons in name paragraphs with
// nonbreaking spaces
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.appliedParagraphStyle = PAR_3_STYLE;
app.findGrepPreferences.findWhat = '(?<!;) ';
app.changeGrepPreferences.changeTo = '~S';
doc.changeGrep();
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// loop over all of the text frames in the doc (there should be just one text
// frame per page)
for (var i = 0; i < doc.pages.count() - 1; i++) {

  var textFrame = doc.pages[i].textFrames.firstItem();
  
  // if the first paragraph of the text frame is a spacing paragraph, remove it
  var firstPar = textFrame.paragraphs.firstItem();
  if (firstPar.appliedParagraphStyle == PAR_4_STYLE) {
    firstPar.remove();
  }

  // if the last paragraph of the text frame is *not* a name line or a spacing
  // paragraph, push this group to the next page with a frame break
  var lastPar = textFrame.paragraphs.lastItem();

  if (!lastPar.validLastPar(textFrame)) {

    var currentPar = lastPar;
    var prevPar = textFrame.paragraphs.previousItem(lastPar);

    while (true) {

      if (prevPar.appliedParagraphStyle == PAR_4_STYLE) {
        try {
          var insertionPoint = currentPar.insertionPoints.firstItem();
          insertionPoint.contents = SpecialCharacters.FRAME_BREAK;
          insertionPoint.applyParagraphStyle(PAR_4_STYLE);
        } catch (err) {
          boop(currentPar.contents);
        }

        // add another page, if we've added overflow
        var lastPage = doc.pages.lastItem();
        var lastTextFrame = lastPage.textFrames.firstItem();
        if (lastTextFrame.overflows) {
          var newLastPage = doc.pages.add();
          var newLastTextFrame = makeTextFrame(newLastPage);
          lastTextFrame.nextTextFrame = newLastTextFrame;
        }

        break;
      }

      currentPar = prevPar;
      prevPar = textFrame.paragraphs.previousItem(currentPar);

    }
  }
}

