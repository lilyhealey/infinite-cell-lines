// ***************************************************************************
// infinite-cell-lines.js
// ***************************************************************************

// needs to be in same folder as script
var TEMPLATE_FILE = 'template.indt';
var DATA_FILE = 'data.tsv';
var DATA_LIMIT = 100;
var POINT_INCREMENT = 0.25;
var MAX_POINT_SIZE = 48;

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

  var scriptFile = app.activeScript;
  return scriptFile.parent.fsName;

}

function getTemplateFile() {
  var templatePath = getScriptFolder() + '/' + TEMPLATE_FILE;
  return new File(templatePath);
}

function getDataFile() {
  var dataPath = getScriptFolder() + '/' + DATA_FILE;
  return new File(dataPath);
}

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

for (var i = 0; i < data.length; i++) {

  var age = data[i].age;
  var population = data[i].population;
  var sex = data[i].sex;
  var disease = data[i].disease;
  var name = data[i].name;
  var synonyms = data[i].synonyms;
  var tissueOfOrigin = data[i].tissueOfOrigin;

  var line1 = '';
  var line2 = '';
  var line3 = '';

  var numLinesToCheck = 1;
  var useLine2 = true;

  // compute line 1
  if (age || population || sex) {

    numLinesToCheck++;

    line1 += age;

    if (age && population) {
      line1 += ' '
    }

    line1 += population;

    if (population && sex || age && sex) {
      line1 += ' ';
    }

    line1 += sex;
  } else {
    line1 = null;
  }

  // compute line 2
  if (disease) {
    numLinesToCheck++;
    line2 += 'with ' + disease;
  } else {
    line2 = null;
    useLine2 = false;
  }

  // compute line 3
  if (name || synonyms || tissueOfOrigin) {

    numLinesToCheck++;

    line3 += name;

    if (name && synonyms) {
      line3 += '; ';
    }

    line3 += synonyms;

    if (synonyms && tissueOfOrigin || name && tissueOfOrigin) {
      line3 += '; ';
    }

    line3 += tissueOfOrigin;
  } else {
    line3 = null;
  }

  // add line 1 to text frame (if it exists)
  if (line1) {
    var insertionPoint = textFrame.insertionPoints.lastItem();
    insertionPoint.contents += line1 + '\r';
    insertionPoint.appliedParagraphStyle = LINE_1_STYLE;
  }

  // add line 2 to text frame (if it exists)
  if (line2) {
    var insertionPoint = textFrame.insertionPoints.lastItem();
    insertionPoint.contents += line2 + '\r';
    insertionPoint.appliedParagraphStyle = LINE_2_STYLE;
  }

  // add line 3 to text frame (if it exists)
  if (line3) {
    var insertionPoint = textFrame.insertionPoints.lastItem();
    insertionPoint.contents += line3 + '\r';
    insertionPoint.appliedParagraphStyle = LINE_3_STYLE;
  }

  // for (var j = 0; j < textFrame.lines.length; j++) {
  //   var line = textFrame.lines[j];
  //   if (line.appliedParagraphStyle == doc.paragraphStyles.item('disease')) {
  //     var contents = line.contents;
  //     for (var k = 0; k < 100; k++) {
  //       if (contents != textFrame.lines[j].contents) {
  //         textFrame.lines[j].textStyleRanges[0].pointSize -= POINT_INCREMENT;
  //         textFrame.lines[j-1].textStyleRanges[0].pointSize = textFrame.lines[j].textStyleRanges[0].pointSize;
  //         break;
  //       }
  //       textFrame.lines[j].textStyleRanges[0].pointSize += POINT_INCREMENT;
  //     }
  //   }
  // }

  var lines = textFrame.lines;
  for (var j = 1; j <= numLinesToCheck; j++) {
    var index = lines.length - j;
    var line = lines[index];
    if (useLine2) {
      if (line.appliedParagraphStyle == LINE_2_STYLE) {
        var contents = line.contents;
        while (true) {
          if (contents != textFrame.lines[index].contents) {
            line.textStyleRanges[0].pointSize -= POINT_INCREMENT;
            lines[index-1].textStyleRanges[0].pointSize = line.textStyleRanges[0].pointSize;
            break;
          }
          line.textStyleRanges[0].pointSize += POINT_INCREMENT;
          if (line.textStyleRanges[0].pointSize >= MAX_POINT_SIZE) {
            break;
          }
        }
      }
    } else {
      if (line.appliedParagraphStyle == LINE_1_STYLE) {
        var contents = line.contents;
        while (true) {
          if (contents != textFrame.lines[index].contents) {
            line.textStyleRanges[0].pointSize -= POINT_INCREMENT;
            break;
          }
          line.textStyleRanges[0].pointSize += POINT_INCREMENT;
          if (line.textStyleRanges[0].pointSize >= MAX_POINT_SIZE) {
            break;
          }
        }
      }
    }
  }

  var insertionPoint = textFrame.insertionPoints.lastItem();
  insertionPoint.contents += '\r';
  // insertionPoint.appliedParagraphStyle = 'space between';
  insertionPoint.appliedParagraphStyle = LINE_4_STYLE;
  try {
    boop(insertionPoint.endBaseline);
  } catch (err) {
    boop('oops');
  }

  if (currentPage.bounds[2] - currentPage.marginPreferences.bottom - insertionPoint.endBaseline < 300) {
    var lastPage = doc.pages.lastItem();
    if (currentPage == lastPage) {
      currentPage = doc.pages.add();
      boop('making a new page');
    } else {
      currentPage = lastPage;
    }

    textFrame = makeTextFrame(currentPage);
    // var nextTextFrame = makeTextFrame(currentPage);
    // textFrame.nextTextFrame = nextTextFrame;
    // var iPoint = textFrame.insertionPoints.lastItem();
    // iPoint.contents = SpecialCharacters.FRAME_BREAK;
    // textFrame = nextTextFrame;
  }
}
