const path = require('path');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');

const { DocParser } = require('./regexFunctions');
const textract = require('textract');

const inputWorkbook = 'test-sheet-og.xlsx';
const nameOfModifiedWorkbook = 'parsedTable';

let isCombinedDoc = true;

const directoryPath = './cleaned-combined-docs';

executeParsing(directoryPath, inputWorkbook);

function executeParsing(directoryPath, inputWorkbook) {
  // Load an Excel Table and process each file from a directory

  fs.readdir(directoryPath, (err, files) => {
    if (err) {
      return console.log('Unable to scan directory: ' + err);
    }
    if (path.extname(inputWorkbook) != '.xlsx') {
      return console.log(
        'Невалидно разширение. Моля, използвайте .xlsx таблици!'
      );
    } else {
      let workbook = XlsxPopulate.fromFileAsync(inputWorkbook);
      let rowNum = workbook.then((workbook) => {
        return findFreeRow(workbook, 1);
      });
      Promise.all([workbook, rowNum]).then((vals) => {
        let workbook = vals[0];
        let rowNum = vals[1];
        processAllFiles(files, workbook, rowNum);
      });
    }
  });
}

function findFreeRow(workbook, rowNum) {
  while (workbook.sheet('Sheet1').row(rowNum).cell(1).value() != undefined) {
    rowNum++;
  }
  return rowNum;
}

// Create a promise for each file in the array, parse the file and return the promise. Once all of them are resolved,
// update the workbook

function processAllFiles(files, workbook, rowNum) {
  // files = ['demo2.doc'];
  // files = ['2772_25-07-2016.doc']; // Debugging one file at a time

  if (isCombinedDoc) {
    files.forEach((doc) => {
      if (path.extname(doc) == '.doc' || path.extname(doc) == '.docx') {
        let docPath = path.join(directoryPath, doc);
        docTexts = splitTexts(docPath);

        docTexts.then((docTexts) => {
          docTexts = Array.from(docTexts);
          processEachFile(docTexts, workbook, rowNum);
        });
      } else {
        console.log(
          `Use a valid .doc or .docx file. Not: ${path.extname(doc)}`
        );
      }
    });
  } else {
    processEachFile(files, workbook, rowNum);
  }
}

function processEachFile(files, workbook, rowNum) {
  let requests = files.reduce((promiseChain, doc) => {
    return promiseChain.then(() =>
      new Promise((resolve) => {
        if (isCombinedDoc) {
          let docPath = 'combinedFile';

          parseFile(doc, docPath, workbook, rowNum, resolve);
        } else if (
          path.extname(doc) == '.doc' ||
          path.extname(doc) == '.docx'
        ) {
          let docPath = path.join(directoryPath, doc);
          //TODO Add error handling for invalid files
          docText = extractText(docPath);
          docText.then((docText) => {
            parseFile(docText, docPath, workbook, rowNum, resolve);
          });
        } else {
          console.log(
            `Use a valid .doc or .docx file. Not: ${path.extname(doc)}`
          );
        }
      })
        .then((lastUsedRow) => {
          // console.log(doc + ' processed.');
          rowNum = lastUsedRow;
        })
        .catch((err) => {
          console.log('Error parsing the file\n' + err);
        })
    );
  }, Promise.resolve());

  requests.then(() => updateWorkbook(workbook));
}

function parseFile(doc, docPath, workbook, rowNum, resolve) {
  let parser = new DocParser(doc, docPath, workbook, rowNum);
  try {
    parser.parseAllFields();
    let lastUsedRow = parser.returnLastUsedRow();

    resolve(lastUsedRow);
  } catch (err) {
    console.log(
      `Could not parse field, due to:\n ${err}\nDoc name: ${docPath}`
    );
  }
}

function updateWorkbook(workbook) {
  let savingDir = path.dirname(inputWorkbook);
  let savingName = nameOfModifiedWorkbook;
  let savingExt = path.extname(inputWorkbook);
  let savingNameExt = savingName.concat(savingExt);
  updatedWorkbook = path.join(savingDir, savingNameExt);
  workbook.toFileAsync(updatedWorkbook);

  let message = `\n✅ Данните са запазени в таблица: ${savingNameExt}`;
  console.log(message);

  moveOutputFile(updatedWorkbook);
}

function moveOutputFile(updatedWorkbook) {
  let source = fs.createReadStream(updatedWorkbook);
  let dest = fs.createWriteStream(
    path.join('/Users/michelangelo/Desktop', updatedWorkbook)
  );

  source.pipe(dest);
  source.on('end', function () {
    console.log('copied!');
  });
  source.on('error', function (err) {
    console.log(err);
  });
}

function extractText(filePath) {
  return new Promise((resolve) => {
    options = { preserveLineBreaks: true };
    textract.fromFileWithPath(filePath, options, (err, result) => {
      if (err) {
        console.log(err);
      }
      resolve(result);
    });
  });
}

function splitTexts(filePath) {
  return new Promise((resolve) => {
    reSplitTexts = /([\s\S\n]*?)(?:\n\n\n[\dЗ][\/\d]?[\dЗ]\n)/g;
    docText = extractText(filePath);
    docText.then((result) => {
      results = [...result.matchAll(reSplitTexts)];
      texts = results.map((text) => text[1]);
      resolve(texts);
    });
  });
}
