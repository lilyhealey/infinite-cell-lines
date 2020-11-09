// ***************************************************************************
// infinite-cell-lines.js
// ***************************************************************************

// needs to be in same folder as script
var TEMPLATE_FILE = 'template.indt';
var DATA_FILE = 'data.tsv';
var DATA_LIMIT = 200;
// var DATA_LIMIT = 10;
var POINT_INCREMENT = 0.25;
var MAX_POINT_SIZE = 48;

/**
 * Logs a string to ~/Desktop/Logs/indesign_log.txt
 * 
 * @param {string} s - the string to log
 */
function boop(s) {

  var now = new Date();
  var output = now.toLocaleString() + ': ' + s;
  // var output = s;
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
var LINE_1_STYLE = doc.paragraphStyles.item('age-population-sex');
var LINE_2_STYLE = doc.paragraphStyles.item('disease');
var LINE_3_STYLE = doc.paragraphStyles.item('names');
var LINE_4_STYLE = doc.paragraphStyles.item('space between');

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

  // add a new text frame
  var nextTextFrame = makeTextFrame(currentPage);
  textFrame.nextTextFrame = nextTextFrame;
  // parInsertionPoint.contents = SpecialCharacters.FRAME_BREAK;
  // parInsertionPoint.appliedParagraphStyle = LINE_4_STYLE;
  // var emptyPar = textFrame.paragraphs.previousItem(textFrame.paragraphs.lastItem());
  // if (!(/\w/.test(emptyPar.contents))) {
  //   emptyPar.remove();
  // }

  textFrame = nextTextFrame;

}

for (var i = 0; i < parGroups.length; i++) {

  var par1 = parGroups[i].par1;
  var par2 = parGroups[i].par2;
  var par3 = parGroups[i].par3;

  var numParsToAdd = 0;
  var usePar2 = true;

  // add paragraph 1 to text frame (if it exists)
  if (par1) {
    
    // add the text and style it
    var insertionPoint = story.insertionPoints.lastItem();
    insertionPoint.contents += par1 + '\r';
    insertionPoint.appliedParagraphStyle = LINE_1_STYLE;

    // make sure to check this paragraph
    numParsToAdd++;
  }

  // add paragraph 2 to text frame (if it exists)
  if (par2) {

    // add the text and style it
    var insertionPoint = story.insertionPoints.lastItem();
    insertionPoint.contents += par2 + '\r';
    insertionPoint.appliedParagraphStyle = LINE_2_STYLE;

    // figure out how many lines we've added
    numParsToAdd++;
  } else {
    usePar2 = false;
  }

  // add paragraph 3 to text frame (if it exists)
  if (par3) {
    
    // add the text and style it
    var insertionPoint = story.insertionPoints.lastItem();
    insertionPoint.contents += par3 + '\r';
    insertionPoint.appliedParagraphStyle = LINE_3_STYLE;

    // figure out how many lines we've added
    numParsToAdd++;
  }

  var pars = story.paragraphs;
  var firstAddedPar = story.paragraphs[story.paragraphs.count() - numParsToAdd];
  var parInsertionPoint = firstAddedPar.insertionPoints[0];

  for (var j = 1; j <= numParsToAdd; j++) {

    var index;
    var par;

    if (textFrame.overflows) {
      addNewTextFrame();
      index = story.paragraphs.length - j - 1;
    } else {
      index = story.paragraphs.length - j;
    }

    par = story.paragraphs[index];

    boop(par.lines.count() + ' - ' + par.contents);

    if (usePar2) {
      if (par.appliedParagraphStyle == LINE_2_STYLE) {
        while (true) {
          if (par.lines.count() > 1) {
            par.textStyleRanges[0].pointSize -= POINT_INCREMENT;
            pars[index - 1].textStyleRanges[0].pointSize = par.textStyleRanges[0].pointSize;
            break;
          } else {
            par.textStyleRanges[0].pointSize += POINT_INCREMENT;
            if (par.textStyleRanges[0].pointSize >= MAX_POINT_SIZE) {
              break;
            }
          }
        }
      }
    } else {
      if (par.appliedParagraphStyle == LINE_1_STYLE) {
        while (true) {
          if (par.lines.count() > 1) {
            par.textStyleRanges[0].pointSize -= POINT_INCREMENT;
            break;
          }
          par.textStyleRanges[0].pointSize += POINT_INCREMENT;
          if (par.textStyleRanges[0].pointSize >= MAX_POINT_SIZE) {
            break;
          }
        }
      }
    }

    boop(par.lines.count() + ' - ' + par.contents);
  }

  if (textFrame.overflows) {
    addNewTextFrame();
  }

  var lastInsertionPoint = story.insertionPoints.lastItem();
  lastInsertionPoint.contents += '\r';
  lastInsertionPoint.appliedParagraphStyle = LINE_4_STYLE;

  if (textFrame.overflows) {
    addNewTextFrame();
  }

}
