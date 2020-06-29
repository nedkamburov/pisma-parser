const path = require('path');

/// Scrape the information from the input file
class DocParser {
  constructor(doc, docPath, workbook, rowNum) {
    this.doc = doc;
    this.docName = path.basename(docPath);
    this.workbook = workbook;
    this.rowNum = rowNum;
    this.lastUsedRow;
  }

  getAllEstates() {
    let reAbout = /(?:Инвестиционно предложение:|Относно\s*:).+\n?.+/i;
    let reEstatesNums = /[№(?:No)]\s*(\d{6}\d*)|(?:Идентификатор)?\s+№*\s*(\d{5}\.\s*\d+\.\s*\d+)|(пл\.*\s№\d+[^,](?:,\s*кв.\s*\d+)*)|(УПИ\s*\D*\d+\-*\d+(?:(?:,\s*кв\.\s*\d+)|(?:\s*[„"'].[^"'“]+["'“]))?(?:\s*в\s*кв\.?\s*\d+)?)/gi;
    //TODO: УПИ шаблонът не засича 'кв.\d'
    let reEstatesNumsCleanUp = /(?:идентификатор\s*)|^(?:\s+|№)|(?:пл\.?\s*№?)/i;

    // let estates = [];
    let counter = 0;
    let searchArea = this.doc.match(reAbout);
    if (searchArea == null) {
      searchArea = this.doc;
    } else {
      searchArea = searchArea[0];
    }

    // Unit-testing for the parsed area of the document
    // console.log(searchArea + '\n\n --------------- \n\n');

    let result = searchArea.match(reEstatesNums);
    let cell1 = this.workbook.sheet('Sheet1').row(this.rowNum).cell('A');
    let cell2 = this.workbook.sheet('Sheet1').row(this.rowNum).cell('AA');
    this.lastUsedRow = this.rowNum;

    if (result == null) {
      cell1.value('--');
      cell2.value('--');
      this.lastUsedRow++;
    } else {
      while (result[counter] != null || result[counter] != undefined) {
        let resultParsed = result[counter];
        resultParsed = resultParsed.replace(reEstatesNumsCleanUp, '');
        // console.log(resultParsed);
        // let isValidEKATTE = new RegExp(/\d{5}/).test(resultParsed);
        // if (isValidEKATTE) {
        //   this.getEKATTE(resultParsed);
        // }

        this.workbook
          .sheet('Sheet1')
          .row(this.lastUsedRow)
          .cell('A')
          .value(resultParsed);
        this.workbook
          .sheet('Sheet1')
          .row(this.lastUsedRow)
          .cell('AA')
          .value(resultParsed);
        this.lastUsedRow++;
        counter++;
        // Unit-testing for incrementing each row for each new estate
        // console.log(`New row ${this.lastUsedRow} and num of result ${counter}`);
      }
    }
    return this.lastUsedRow;
  }

  //TODO: Should this get the whole paragraph or risk having cutoff sentences?
  getAbout(newRowNum) {
    let reAbout = /(?:Инвестиционно предложение:|Относно:)(?:(.+)(?=(?:в|на)\s*имот)|[^„'"]+.(.+?(?=[“"'”]|(?=в\s*имот))))/i;
    let result = this.doc.match(reAbout);
    if (result == null) {
      result = '--';
    } else if (result[1] != undefined) {
      result = result[1];
    } else if (result[2] != undefined) {
      result = result[2];
    }

    this.workbook.sheet('Sheet1').row(newRowNum).cell('AI').value(result);
  }

  getMasiv(newRowNum) {
    let reMasiv = /масив\s*(\d+)/i;
    let result = this.doc.match(reMasiv);
    let cell = this.workbook.sheet('Sheet1').row(newRowNum).cell('AB');
    if (result == null) {
      result = '--';
      cell.value(result);
    } else {
      result = result[1];
      cell.value(result);
    }
  }

  getRecipient(newRowNum) {
    let result, resultCompany;
    let reRecipient = /ДО\n(.*)/i;
    let reRecipientCompany = /(?:ЕООД|ООД|ЕТ|ОБЩИНА)/i;
    let reRecipientPerson = /УВАЖАЕМ(?:И|а)\s*(.*(?=\,))/i;
    result = this.doc.match(reRecipient);
    let cell = this.workbook.sheet('Sheet1').row(newRowNum).cell('AM');
    let cellCompany = this.workbook.sheet('Sheet1').row(newRowNum).cell('AJ');

    if (result == null) {
      result = '--';
      resultCompany = '--';
    } else {
      if (reRecipientCompany.test(result[1])) {
        resultCompany = this.titleCase(result[1]);
        result = this.doc.match(reRecipientPerson)[1];
        result = result.replace(/ГОСПОДИН/i, 'г-н');
        result = result.replace(/ГОСПОЖО/i, 'г-жа');
      } else {
        result = result[1];
        resultCompany = result;
      }
    }

    result = this.titleCase(result);
    cell.value(result);
    cellCompany.value(resultCompany);
  }

  getPlace(newRowNum) {
    let rePlace = /ДO*.*\n*(.*)/;
    let result = this.doc.match(rePlace);
    // result = this.titleCase(result[1]);

    // this.workbook.sheet('Sheet1').row(newRowNum).cell('AM').value(result);
  }

  getEKATTE(estateNum, newRowNum) {
    // console.log(estateNum);
  }

  getRIOSV(newRowNum) {
    let reRIOSV = /ДИРЕКТОР\n*\s*НА\s*РИОСВ\s*[-–—]\s*(.+)(?:\:+)?/i;
    let result = this.doc.match(reRIOSV);

    try {
      result = this.titleCase(result[1]);
    } catch {
      result = '--';
    }
    this.workbook.sheet('Sheet1').row(newRowNum).cell('AR').value(result);
  }

  getDocNum(newRowNum) {
    let reDocNum = /Изх\.\s*№(\d+)/i;
    let result = this.doc.match(reDocNum);
    let cell = this.workbook.sheet('Sheet1').row(newRowNum).cell('AX');

    if (result == null) {
      result = '--';
    } else {
      result = result[1];
    }

    cell.value(result);
  }

  getDate(newRowNum) {
    let reDate = /Вх\.*\s*№\s*(?:[\d\s(?:а-яА-Я)-—]*(?=\/)?)?\s*(?:от|\/)?\s*([0-9ОЗ]{2}\s?\.?\s?[0-9ОЗ]{2}\.[0-9ОЗ]{4})\s*(?:г\s*\.?)?/i;
    let result = this.doc.match(reDate);
    let cell = this.workbook.sheet('Sheet1').row(newRowNum).cell('AY');

    if (result == null) {
      result = '--';
    } else {
      result = result[1].replace(/\s/, '');
      result = result.replace(/[зЗ]/, '3');
      result = result.replace(/[оО]/, '0');
      result = `${result} г.`;
    }

    cell.value(result);
  }

  getDecision(newRowNum) {
    let reDecision = /преценката\s*[а-я\s]+,\s*че\s*не\s*е | не\s*следва\s*да\s*бъде\s*провеждана\s*процедур[аи] | не\s*подлежи\s*на\s*процедур[аи] | не\s*е\s*необходимо\s*провеждане\s*на\s*процедур[аи]/i;
    let result = this.doc.match(reDecision);
    let cell = this.workbook.sheet('Sheet1').row(newRowNum).cell('BA');

    if (result == null) {
      cell.value(`Нужен е преглед на ${this.docName}`);
    } else {
      cell.value('Не е необходима процедура');
    }
  }

  getWithinZone(newRowNum) {
    let reWithinZone = /не\sпопада[т]?\sв\sграниците\sна\sзащитена\sтеритория|няма\s*вероятност\s*от\s*отрицателно\s*въздействие/i;
    let result = this.doc.match(reWithinZone);
    let cell = this.workbook.sheet('Sheet1').row(newRowNum).cell('BN');

    if (result == null) {
      cell.value(`Нужен е преглед на ${this.docName}`);
    } else cell.value('Не');
  }

  getZone(newRowNum) {
    let zoneName, zoneNum;
    let reZone = /[“"'„]([а-яА-Я\s]*)['"”]\,?(?:\s*за\s*птиците\,?|\s*за\s*хабитатите\s*\,?)?\s*с?\s*код?\s*BG\s*(\d*)/i;
    let result = this.doc.match(reZone);

    if (result == null) {
      zoneName = '--';
      zoneNum = '--';
    } else {
      zoneName = result[1];
      zoneNum = 'BG' + result[2];

      // Unit-testing for the zone retrieval
      // result.input = '';
      // console.log(result);
    }

    this.workbook.sheet('Sheet1').row(newRowNum).cell('CR').value(zoneName);
    this.workbook.sheet('Sheet1').row(newRowNum).cell('CS').value(zoneNum);
  }

  setStaticFields(newRowNum) {
    this.workbook.sheet('Sheet1').row(newRowNum).cell('B').value('Да');
    this.workbook.sheet('Sheet1').row(newRowNum).cell('AQ').value('РИОСВ');
    this.workbook.sheet('Sheet1').row(newRowNum).cell('AU').value('Писмо');
    this.workbook
      .sheet('Sheet1')
      .row(newRowNum)
      .cell('AZ')
      .value('Писмо по чл.2, ал.2 от Наредбата за ОС');
  }

  titleCase(str) {
    let splitStr = str.toLowerCase().split(' ');

    for (let i = 0; i < splitStr.length; i++) {
      splitStr[i] =
        splitStr[i].charAt(0).toUpperCase() + splitStr[i].substring(1);
    }
    return splitStr.join(' ');
  }

  returnLastUsedRow() {
    return this.lastUsedRow;
  }

  parseAllFields() {
    try {
      this.getAllEstates(); // Skips very messy estate numbers for now...

      for (let i = this.rowNum; i < this.lastUsedRow; i++) {
        let newRowNum = i;

        this.getAbout(newRowNum);
        this.getMasiv(newRowNum);
        this.getRecipient(newRowNum);
        this.getRIOSV(newRowNum);
        this.getDocNum(newRowNum); // The input is handwritten, hard to process...
        this.getDate(newRowNum);
        // this.getPlace(newRowNum);
        this.getDecision(newRowNum);
        this.getWithinZone(newRowNum);
        this.getZone(newRowNum); // Doesn't get multiple zones
        this.setStaticFields(newRowNum);
      }
    } catch (err) {
      console.log(err);
    }
  }
}

exports.DocParser = DocParser;
