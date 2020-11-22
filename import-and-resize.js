// ***************************************************************************
// import-and-resize.js
// ***************************************************************************

// these two files need to be in the same folder as the script
var TEMPLATE_FILE = 'template.indt';
var DATA_FILE = 'data.tsv';

// limit the number of rows of data to read (for testing); set to null to read
// all data
var DATA_LIMIT = 500;

// resizing point increment; the smaller this number, the more precise the
// justification will be, but the slower the script will run
var POINT_INCREMENT = .25;

// maximum point size when resizing
var MAX_POINT_SIZE = 48;

// not currently used (set minimum via the paragraph style in the template
// instead)
var MIN_POINT_SIZE = 16;

// space between text frames, in points
var SPACE_BETWEEN_FRAMES = 15;

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

function getScriptFolder() {
  return app.activeScript.parent.fsName;
}

function getTemplateFile() {
  var templatePath = getScriptFolder() + '/' + TEMPLATE_FILE;
  return new File(templatePath);
}

function getDataFile() {
  var dataPath = getScriptFolder() + '/' + DATA_FILE;
  return new File(dataPath);
}

/**
 * Reads tab-separated values into a javascript object.
 * 
 * @param {File} dataFile - the file with tsv data
 * @returns{Array[Object]}
 */
function readData(dataFile) {

  var data = [];
  var line;

  if (dataFile.open('r')) {

    // get rid of the headers
    line = dataFile.readln();

    while (true) {
      line = dataFile.readln();

      if (line) {

        var parts = line.split('\t');
        data.push({
          age: parts[0],
          population: parts[1],
          sex: parts[2],
          disease: parts[3],
          name: parts[4],
          synonyms: parts[5],
          tissueOfOrigin: parts[6],
        });

        if (DATA_LIMIT && data.length >= DATA_LIMIT) {
          break;
        }
      } else {
        // no more lines left in the file
        break;
      }
    }
  } else {
    alert('something went wrong');
  }

  return data;
}

/**
 * Converts data to paragraphs.
 * 
 * @param {*} data 
 */
function dataToParagraphs(data) {

  var parGroups = [];

  for (var i = 0; i < data.length; i++) {

    var age = data[i].age;
    var population = data[i].population;
    var sex = data[i].sex;
    var disease = data[i].disease;
    var name = data[i].name;
    var synonyms = data[i].synonyms;
    var tissueOfOrigin = data[i].tissueOfOrigin;

    var par1 = '';
    var par2 = '';
    var par3 = '';

    // compute par 1
    if (age || population || sex) {

      par1 += age;

      if (age && population) {
        par1 += ' '
      }

      par1 += population;

      if (population && sex || age && sex) {
        par1 += ' ';
      }

      par1 += sex;
    } else {
      par1 = null;
    }

    // compute line 2
    if (disease) {
      par2 += 'with ' + disease;
    } else {
      par2 = null;
    }

    // compute line 3
    if (name || synonyms || tissueOfOrigin) {

      par3 += name;

      if (name && synonyms) {
        par3 += '; ';
      }

      par3 += synonyms;

      if (synonyms && tissueOfOrigin || name && tissueOfOrigin) {
        par3 += '; ';
      }

      par3 += tissueOfOrigin;
    } else {
      par3 = null;
    }

    parGroups.push({
      par1: par1,
      par2: par2,
      par3: par3
    });
  }

  return parGroups;
}

/**
 * Make a text frame that fills the page margins.
 * 
 * @param {Page} page - the page to make the text frame on
 * @returns {TextFrame} - the created text frame
 */
function makeTextFrame(page, y1) {

  if (!y1) {
    y1 = page.marginPreferences.top;
  }
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

  var textFrame = page.textFrames.add(
    undefined,
    undefined,
    undefined,
    {
      geometricBounds: [
        y1, x1, y2, x2
      ]
    }
  );

  // make sure the top of the caps of the first line align with the top of the
  // text frame
  textFrame.textFramePreferences.firstBaselineOffset = FirstBaseline.CAP_HEIGHT;

  return textFrame;
}

// read in the data
var tsvFile = getDataFile();
var data = readData(tsvFile);
var parGroups = dataToParagraphs(data);

// open the template doc
var doc = app.open(getTemplateFile());

// make sure we're doing everything in points
doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;

// get paragraph styles we're going to use
var PAR_1_STYLE = doc.paragraphStyles.item('age-population-sex');
var PAR_2_STYLE = doc.paragraphStyles.item('disease');
var PAR_3_STYLE = doc.paragraphStyles.item('names');
var PAR_4_STYLE = doc.paragraphStyles.item('space between');

// make the initial text frame
var currentPage = doc.pages.firstItem();
var textFrame = makeTextFrame(currentPage);
var story = textFrame.parentStory;

function addNewTextFrame() {

  // add a new page if necessary
  var lastPage = doc.pages.lastItem();
  if (currentPage == lastPage) {
    currentPage = doc.pages.add();
  } else {
    currentPage = lastPage;
  }

  // this is a little hacky, but it works:
  // add a new text frame to the next page
  var nextTextFrame = makeTextFrame(currentPage);
  // move the current text frame to the new page
  textFrame.move(currentPage);
  textFrame.geometricBounds = nextTextFrame.geometricBounds;
  // remove the interim text frame
  nextTextFrame.remove();
}

for (var i = 0; i < parGroups.length; i++) {

  // {age} {population} {sex}
  var par1 = parGroups[i].par1;

  // {disease}
  var par2 = parGroups[i].par2;

  // {name}; {synonyms}; {tissueOfOrigin}
  var par3 = parGroups[i].par3;

  // add paragraph 1 to text frame (if it exists)
  if (par1) {
    // add the text and style it
    var insertionPoint = story.insertionPoints.lastItem();
    insertionPoint.contents += par1 + '\r';
    insertionPoint.appliedParagraphStyle = PAR_1_STYLE;
  }

  // add paragraph 2 to text frame (if it exists)
  if (par2) {
    // add the text and style it
    var insertionPoint = story.insertionPoints.lastItem();
    insertionPoint.contents += par2 + '\r';
    insertionPoint.appliedParagraphStyle = PAR_2_STYLE;
  }

  // add paragraph 3 to text frame (if it exists)
  if (par3) {
    // add the text and style it
    var insertionPoint = story.insertionPoints.lastItem();
    insertionPoint.contents += par3 + '\r';
    insertionPoint.appliedParagraphStyle = PAR_3_STYLE;
  }

  if (textFrame.overflows) {
    addNewTextFrame();
  }

  // resize the paragraphs as necessary
  var numPars = story.paragraphs.count();

  // figure out the number of lines in the first two paragraphs
  var par1NumLines = 0;
  var par2NumLines = 0;
  var longestLine = story.lines.firstItem();
  var allLines = [];

  for (var j = 0; j < story.lines.count(); j++) {

    var line = story.lines[j];

    if (line.appliedParagraphStyle == PAR_1_STYLE) {
      allLines.push(line);
      if (line.endHorizontalOffset > longestLine.endHorizontalOffset) {
        longestLine = line;
      }
      par1NumLines++;
      continue;
    } else if (line.appliedParagraphStyle == PAR_2_STYLE) {
      allLines.push(line);
      if (line.endHorizontalOffset > longestLine.endHorizontalOffset) {
        longestLine = line;
      }
      par2NumLines++;
    }
  }

  if (par1NumLines > 1 || par2NumLines > 1) {

    var lineTSR = longestLine.textStyleRanges[0];
    var contents = longestLine.contents;

    while (longestLine.lines.count() < 2 && !textFrame.overflows) {

      lineTSR.pointSize += POINT_INCREMENT;

      if (lineTSR.pointSize > MAX_POINT_SIZE) {
        break;
      }

    }
    lineTSR.pointSize -= POINT_INCREMENT;

    for (var k = 0; k < allLines.length; k++) {
      var line = allLines[k];
      line.textStyleRanges[0].pointSize  = lineTSR.pointSize;
    }

  } else {
    for (var j = 0; j < numPars; j++) {
  
      // if (textFrame.overflows) {
      //   addNewTextFrame();
      // }
  
      var par = story.paragraphs[j];
      var parTSR = par.textStyleRanges[0];
      var prevPar = j > 0 ? story.paragraphs.previousItem(par) : null;
      var prevParTSR = prevPar ? prevPar.textStyleRanges[0] : null;
  
      // if there is a disease paragraph, use it for sizing; otherwise, use
      // demographic paragraph
      if (par2) {
  
        // if this paragraph we're checking is the disease paragraph, resize it
        if (par.appliedParagraphStyle == PAR_2_STYLE) {
  
          // increase the type size by POINT_INCREMENT until the paragraph spans
          // more than one line
          while (true) {
  
            if (par.lines.count() > 1) {
  
              // decrease the paragraph by POINT_INCREMENT if it is not the
              // minimmum
              if (parTSR.pointSize > parseInt(PAR_2_STYLE.pointSize)) {
                parTSR.pointSize -= POINT_INCREMENT;
              }
  
              // resize the previous paragraph, if it is demographic info
              if (prevPar && prevPar.appliedParagraphStyle == PAR_1_STYLE) {
                prevParTSR.pointSize = parTSR.pointSize;
              }
  
              break;
  
            } else {
  
              parTSR.pointSize += POINT_INCREMENT;
  
              if (parTSR.pointSize >= MAX_POINT_SIZE) {
  
                // resize the previous paragraph, if necessary
                if (prevPar && prevPar.appliedParagraphStyle == PAR_1_STYLE) {
                  prevParTSR.pointSize = parTSR.pointSize;
                }
  
                break;
              }
            }
          }
  
          // if, by resizing the disease paragraph, we've pushed the demographic
          // paragraph to two lines, reduce the size of the demographic paragraph
          // until it fits one line
          if (
            prevPar &&
            prevPar.appliedParagraphStyle == PAR_1_STYLE &&
            prevPar.lines.count() > 1
          ) {
            while (true) {
              if (prevPar.lines.count() == 1) {
                parTSR.pointSize = prevParTSR.pointSize;
                break;
              }
              prevParTSR.pointSize -= POINT_INCREMENT;
            }
          }
        }
      } else {
        // no disease paragraph; use demographic paragraph for sizing
        if (par.appliedParagraphStyle == PAR_1_STYLE) {
          while (true) {
            if (par.lines.count() > 1) {
              parTSR.pointSize -= POINT_INCREMENT;
              break;
            }
            parTSR.pointSize += POINT_INCREMENT;
            if (parTSR.pointSize >= MAX_POINT_SIZE) {
              break;
            }
          }
        }
      }
    }
  }


  if (textFrame.overflows) {
    addNewTextFrame();
  }

  // resize the current text frame to be only as tall as it needs to be
  var lastLine = textFrame.lines.lastItem();
  var gb = textFrame.geometricBounds;
  textFrame.geometricBounds = [
    gb[0],
    gb[1],
    lastLine.endBaseline,
    gb[3]
  ];

  // add a new text frame for the next row
  textFrame = makeTextFrame(
    currentPage, 
    lastLine.endBaseline + SPACE_BETWEEN_FRAMES
  );
  story = textFrame.parentStory;

}

// remove the extra text frame created in the last iteration of the loop
textFrame.remove();
