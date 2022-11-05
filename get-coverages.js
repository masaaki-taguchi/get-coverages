'use strict';

const jsforce = require('jsforce');
const xlsx = require('xlsx-populate');

const DEFAULT_USER_CONFIG_PATH = './user_config.yaml';
const COMMAND_OPTION_SILENT_MODE = '-s';
const COMMAND_OPTION_INCLUDE_INVALID_COVERAGE = '-i';
const COMMAND_OPTION_HELP = '-h';
const COMMAND_OPTION_DISPLAY_TO_CONSOLE = '-d';
const COMMAND_OPTION_EXECUTE_APEX_TEST = '-t';
const COMMAND_OPTION_CONFIG = '-c';
const TEST_LEVEL_ALL_TEST = 'RunLocalTests';
const APEX_TEST_JOB_CHECK_INTERVAL = 3000;

const DATE_POSITION = [1, 8];
const WARNING_HEADER_POSITION = [2, 7];
const FATAL_HEADER_POSITION = [2, 8];
const RESULT_START_POSITION_Y = 3;
const RESULT_INDEX_POSITION_X = 1;
const RESULT_APEX_CLASS_NAME_POSITION_X = 2;
const RESULT_COVERGE_POSITION_X = 3;
const RESULT_TOTAL_LINES_POSITION_X = 4;
const RESULT_CONVERD_LINES_POSITION_X = 5;
const RESULT_UNCONVERD_LINES_POSITION_X = 6;
const RESULT_WARNING_MARK_POSITION_X = 7;
const RESULT_FATAL_MARK_POSITION_X = 8;

// default parameter
global.enabledLogging = true;
let userConfigPath = DEFAULT_USER_CONFIG_PATH;
let enabledInvalidCoverage = false;
let enabledDisplayToConsole = false;
let isSelectedExecuteApexTest = false;

// analyzes command line options
let paramList = [];
let paramCnt = 0;
for (let i = 2; i < process.argv.length; i++) {
  if (process.argv[i] === COMMAND_OPTION_SILENT_MODE) {
    global.enabledLogging = false;
  } else if (process.argv[i] === COMMAND_OPTION_INCLUDE_INVALID_COVERAGE) {
    enabledInvalidCoverage = true;
  } else if (process.argv[i] === COMMAND_OPTION_DISPLAY_TO_CONSOLE) {
    enabledDisplayToConsole = true;
  } else if (process.argv[i] === COMMAND_OPTION_EXECUTE_APEX_TEST) {
    isSelectedExecuteApexTest = true;
  } else if (process.argv[i] === COMMAND_OPTION_CONFIG) {
    if (i + 1 >= process.argv.length) {
      usage();
    }
    userConfigPath = process.argv[i + 1];
  } else if (process.argv[i] === COMMAND_OPTION_HELP) {
    usage();
  } else {
    paramList.push(process.argv[i]);
  }
}

loadUserConfig(userConfigPath);
let userConfig = global.userConfig;
let conn = new jsforce.Connection({loginUrl: userConfig.loginUrl, version: userConfig.apiVersion});

(async () => {
  // login to salesforce
  logging('Settings:');
  logging('  loginUrl:' + userConfig.loginUrl);
  logging('  apiVersion:' + userConfig.apiVersion);
  logging('  userName:' + userConfig.userName);
  logging('  templateFilePath:' + userConfig.templateFilePath);
  logging('  resultFilePath:' + userConfig.resultFilePath);
  logging('  warningPercent:' + userConfig.warningPercent);
  logging('  fatalPercent:' + userConfig.fatalPercent);
  logging('  targetApexClass:' + userConfig.targetApexClass);
  logging('  targetApexTestClass:' + userConfig.targetApexTestClass);

  await conn.login(userConfig.userName, userConfig.password, function (err, userInfo) {
    if (err) {
      console.error(err);
      process.exit(1);
    }
  });
  if (isSelectedExecuteApexTest) {
    executeApexTest();
  } else {
    getCoverage();
  }
})();

function executeApexTest() {
  (async () => {
    try {
      let requestJson = null;
      if (global.targetApexTestClassArray.length !== 0) {
        let requestBodyJson = [];
        global.targetApexTestClassArray.forEach((val) => {
          requestBodyJson.push({className: val});
        });
        requestJson = {"tests": requestBodyJson};
      } else {
        requestJson = {"testLevel": TEST_LEVEL_ALL_TEST};
      }
      let testJobId = null;
      await conn.requestPost('/tooling/runTestsAsynchronous/', requestJson, undefined, function (err, result) {
        if (err) {
          console.error(err);
          process.exit(1);
        } else if (result) {
          testJobId = result;
        }
      })
      if (testJobId) {
        logging('ApexTestResult:');
        let completeTestSet = new Set();
        let cntOK = 0;
        let cntNG = 0;
        let timerId = setInterval(async function() {
          let recordList = await queryToolingApi(
            'SELECT Id, ApexClass.Name, ExtendedStatus, Status FROM ApexTestQueueItem WHERE ParentJobId = \'' + testJobId + '\''
          );
          let isCompleted = true;
          recordList.forEach((val) => {
            if (val.Status === 'Completed') {
              if (!completeTestSet.has(val.ApexClass.Name)) {
                let matches = /\((\d+)\/(\d+)\)/.exec(val.ExtendedStatus);
                let result = 'NG';
                if (matches && matches[1] === matches[2]) {
                  result = 'OK';
                  cntOK++;
                } else {
                  cntNG++;
                }
                completeTestSet.add(val.ApexClass.Name);
                let progress = Math.floor((completeTestSet.size / recordList.length) * 100) + '%';                
                logging('  [' + progress + ']' + val.ApexClass.Name + ':' + result + val.ExtendedStatus);
              }
            } else {
              isCompleted = false;
            }
          });
          if (isCompleted) {
            clearInterval(timerId);
            logging('  *** ApexTestClassCount:' + completeTestSet.size + ' OK:' + cntOK + ' NG:' + cntNG + ' ***');
            logging('Done.');
          } 
        }, APEX_TEST_JOB_CHECK_INTERVAL);
      }
    } catch(e) {
      console.error(e);
    }
  })();
}

function getCoverage() {
  (async () => {
    let recordList = await queryToolingApi(
      "SELECT ApexClassOrTriggerId, ApexClassOrTrigger.Name, NumLinesCovered, NumLinesUncovered FROM ApexCodeCoverageAggregate"
    );
    let coverageMap = new Map();
    let cnt = 0;
    const targetApexClassSet = global.targetApexClassSet;
    for (let record of recordList) {
      try {
        if (!enabledInvalidCoverage && record.NumLinesCovered === 0 && record.NumLinesUncovered === 0) {
          continue;
        }
        if (targetApexClassSet.size === 0 ||
           (targetApexClassSet.size > 0 && targetApexClassSet.has(record.ApexClassOrTrigger.Name))) {
          let numLinesCovered = record.NumLinesCovered;
          let numLinesUncovered = record.NumLinesUncovered;
          let totalLines = numLinesCovered + numLinesUncovered;
          record.totalLines = totalLines;
          let coverage = numLinesCovered / totalLines;
          if (isNaN(coverage)) {
            coverage = 0;
          }
          record.coverage = Math.floor(coverage * 100) / 100;
          coverageMap.set(record.ApexClassOrTrigger.Name, record);
          cnt++;
        }
      } catch (e) {
        // ignore exception
      }
    }
    logging('ApexClassCount:' + cnt);
    if (cnt === 0) {
      logging('Done.');
      process.exit(0);
    }

    coverageMap = new Map([...coverageMap.entries()].sort(function (a, b) {
        var _a = a[0].toLowerCase();
        var _b = b[0].toLowerCase();
        if (_a < _b) return -1;
        if (_a > _b) return 1;
        return 0;
      })
    );
    if (enabledDisplayToConsole) {
      let values = [];
      for (let key of coverageMap.keys()) {
        let record = coverageMap.get(key);
        values.push({
          ApexClassName: record.ApexClassOrTrigger.Name,
          Coverage: record.coverage * 100 + '%',
          TotalLines: record.totalLines,
          CoveredLines: record.NumLinesCovered,
          UncoveredLines: record.NumLinesUncovered,
        });
      }
      console.table(values);
      logging('Done.');
    } else {
      await xlsx.fromFileAsync(userConfig.templateFilePath).then((workBook) => {
        let xlsxSheet = workBook.sheet(0);
        let templateStyleList = getTemplateStyleList(xlsxSheet);

        let date = new Date();
        let dateStr = getFormattedDateTime(date);
        let xlsxCell = xlsxSheet.row(DATE_POSITION[0]).cell(DATE_POSITION[1]);
        xlsxCell.value(dateStr);

        xlsxCell = xlsxSheet.row(WARNING_HEADER_POSITION[0]).cell(WARNING_HEADER_POSITION[1]);
        let warningPercent = userConfig.warningPercent;
        let warningPercentHeader = '< ' + warningPercent * 100 + '%';
        xlsxCell.value(warningPercentHeader);

        xlsxCell = xlsxSheet.row(FATAL_HEADER_POSITION[0]).cell(FATAL_HEADER_POSITION[1]);

        let fatalPercent = userConfig.fatalPercent;
        let fatalPercentHeader = '< ' + fatalPercent * 100 + '%';
        xlsxCell.value(fatalPercentHeader);

        let resultWorkY = RESULT_START_POSITION_Y;
        let index = 1;

        coverageMap.forEach(function(record, key) {
          let coverage = record.coverage;

          putTemplateStyle(xlsxSheet, templateStyleList, resultWorkY);
          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_INDEX_POSITION_X);
          xlsxCell.value(index);
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_APEX_CLASS_NAME_POSITION_X);
          xlsxCell.value(record.ApexClassOrTrigger.Name);
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_COVERGE_POSITION_X);
          xlsxCell.value(record.coverage);
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_TOTAL_LINES_POSITION_X);
          xlsxCell.value(record.totalLines);
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_CONVERD_LINES_POSITION_X);
          xlsxCell.value(record.NumLinesCovered);
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_UNCONVERD_LINES_POSITION_X);
          xlsxCell.value(record.NumLinesUncovered);
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_WARNING_MARK_POSITION_X);
          if (coverage < warningPercent) {
            xlsxCell.value(userConfig.belowMark);
          }
          fillColorByCoverage(xlsxCell, coverage);

          xlsxCell = xlsxSheet.row(resultWorkY).cell(RESULT_FATAL_MARK_POSITION_X);
          if (coverage < fatalPercent) {
            xlsxCell.value(userConfig.belowMark);
          }
          fillColorByCoverage(xlsxCell, coverage);

          resultWorkY++;
          index++;
        });

        workBook.toFileAsync(userConfig.resultFilePath).then((result) => {});
        logging('Done.');
      });
    }
  })();
}

async function queryToolingApi(query) {
  let recordList = [];
  let nextRecordsUrl = null;
  await conn.tooling.query(query, function (err, result) {
    if (err) {
      console.error(err);
      process.exit(1);
    } else if (result && result.totalSize > 0) {
      recordList = result.records;
      nextRecordsUrl = result.nextRecordsUrl;
    }
  });

  while (nextRecordsUrl) {
    await conn.tooling.queryMore(nextRecordsUrl, function (err, result) {
      nextRecordsUrl = null;
      if (err) {
        console.error(err);
        process.exit(1);
      } else if (result) {
        nextRecordsUrl = result.nextRecordsUrl;
        recordList = recordList.concat(result.records);
      }
    });
  }
  return recordList;
}

function fillColorByCoverage(xlsxCell, coverage) {
  if (coverage < global.userConfig.fatalPercent) {
    xlsxCell.style('fill', global.userConfig.fatalColor);
  } else if (coverage < global.userConfig.warningPercent) {
    xlsxCell.style('fill', global.userConfig.warningColor);
  }
}

function loadYamlFile(fileName) {
  const fs = require('fs');
  let existsFile = fs.existsSync(fileName);
  if (!existsFile) {
    console.error('File not found. filePath: ' + fileName);
    process.exit(1);
  }
  const yaml = require('js-yaml');
  const yamlText = fs.readFileSync(fileName, 'utf8');
  return yaml.load(yamlText);
}

function loadUserConfig(userConfigPath) {
  let userConfigPathWork = userConfigPath;
  if (userConfigPathWork === undefined) {
    userConfigPathWork = DEFAULT_USER_CONFIG_PATH;
  }
  const path = require('path');
  const config = loadYamlFile(path.join(__dirname, userConfigPathWork));
  global.userConfig = config;

  const targetApexClassSet = new Set();
  const userConfigApexClass = config.targetApexClass;
  if (userConfigApexClass !== undefined) {
    for (const i in userConfigApexClass) {
      targetApexClassSet.add(userConfigApexClass[i].trim());
    }
  }
  global.targetApexClassSet = targetApexClassSet;

  const targetApexTestClassArray = new Array();
  const userConfigApexTestClass = config.targetApexTestClass;
  if (userConfigApexTestClass !== undefined) {
    for (const i in userConfigApexTestClass) {
        targetApexTestClassArray.push(userConfigApexTestClass[i].trim());
    }
  }
  global.targetApexTestClassArray = targetApexTestClassArray;
}

function usage() {
  console.log('usage: get-coverages.js [-options]');
  console.log('    -c <pathname> specifies a config file path (default is ./user_config.yaml)');
  console.log('    -s            silent mode');
  console.log('    -i            include invalid coverage');
  console.log('    -d            display test coverage to console');
  console.log('    -t            execute apex test');
  console.log('    -h            display this help');
  process.exit(0);
}

function putTemplateStyle(xlsxSheet, xlsxTemplateStyleList, cellY) {
  for (let i = 0; i < xlsxTemplateStyleList.length; i++) {
    let cell = xlsxSheet.cell(cellY, i + 1);
    cell.style(xlsxTemplateStyleList[i]);
  }
}

function getTemplateStyleList(xlsxSheet) {
  const end_col_num = xlsxSheet.usedRange().endCell().columnNumber();
  let xlsxTemplateStyleList = new Array();
  for (let i = 1; i <= end_col_num; i++) {
    let style = xlsxSheet.cell(3, i).style([
        'bold',
        'italic',
        'underline',
        'strikethrough',
        'subscript',
        'superscript',
        'fontSize',
        'fontFamily',
        'fontColor',
        'horizontalAlignment',
        'justifyLastLine',
        'indent',
        'verticalAlignment',
        'wrapText',
        'shrinkToFit',
        'textDirection',
        'textRotation',
        'angleTextCounterclockwise',
        'angleTextClockwise',
        'rotateTextUp',
        'rotateTextDown',
        'verticalText',
        'fill',
        'border',
        'borderColor',
        'borderStyle',
        'diagonalBorderDirection',
        'numberFormat',
      ]);
    xlsxTemplateStyleList.push(style);
  }

  return xlsxTemplateStyleList;
}

function logging(message) {
  if (global.enabledLogging) {
    const nowDate = new Date();
    console.log('[' + getFormattedDateTime(nowDate) + '] ' + message);
  }
}

function getFormattedDateTime(date) {
  let dateString =
    date.getFullYear() + '/' +
    ('0' + (date.getMonth() + 1)).slice(-2) + '/' +
    ('0' + date.getDate()).slice(-2) + ' ' +
    ('0' + date.getHours()).slice(-2) + ':' +
    ('0' + date.getMinutes()).slice(-2) + ':' +
    ('0' + date.getSeconds()).slice(-2);
  return dateString;
}
