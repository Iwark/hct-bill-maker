var app = require('app');  // Module to control application life.
var BrowserWindow = require('browser-window');  // Module to create native browser window.
var dialog = require('dialog');
var ipc = require('ipc');

var fs = require('fs');
var officegen = require('officegen');
var xlsx = officegen ('xlsx');

var SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

// Report crashes to our server.
require('crash-reporter').start();

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the javascript object is GCed.
var mainWindow = null;

// Quit when all windows are closed.
app.on('window-all-closed', function() {
  if (process.platform != 'darwin')
    app.quit();
});

// This method will be called when atom-shell has done everything
// initialization and ready for creating browser windows.
app.on('ready', function() {
  // Create the browser window.
  mainWindow = new BrowserWindow({width: 800, height: 600});

  xlsx.on('finalize', function(written){
    console.log('Finish to create an Excel File. Totle bytes created: ' + written);
  });

  xlsx.on('error', function(err){
    console.log(err);
  });

  ipc.on('openFileDialog', function(event, arg){
    dialog.showOpenDialog({
      properties: ['openFile']
    }, function(filePath){
      console.log("filePath = " + filePath);

      SpreadsheetReader.read(filePath, function (err, workbook) {
        if(err){
          console.log(err);
          return;
        }
        var writers = [];
        // Iterate on sheets
        workbook.sheets.forEach(function (sheet) {
          console.log('sheet: %s(%d)', sheet.name, sheet.index);
          if(sheet.index != 7) return;
          // Iterate on rows
          var stopLoop = false;
          sheet.rows.forEach(function (row) {
            // Iterate on cells
            row.forEach(function (cell) {
              if(cell.row >= 4 && !stopLoop){
                var writer = {};
                if(writers.length > cell.row - 4) writer = writers[cell.row-4];
                else writers.push(writer);

                if(cell.column == 2){
                  //ライター名
                  writer["name"] = cell.value;
                }
                if(cell.column == 3){
                  //振り込み名
                  if(cell.value) writer["name"] = cell.value;
                }
                if(cell.column == 4){
                  //単価
                  writer["price"] = cell.value;
                }
                if(cell.column == 7){
                  //本数
                  writer["num"] = cell.value;
                  if(cell.value && cell.value > 0) writer["sum"] = writer["price"] * writer["num"];
                }
                if(writer["name"] == "合計") stopLoop = true;
              }
            });
          });
          event.sender.send('asynchronous-reply', writers);
        });

        writers.forEach(function(writer){
          if(writer["num"] && writer["num"] > 0) makeBill(writer);
        });
        var out = fs.createWriteStream('ライター個人請求書.xlsx');

        out.on('error', function(err){
          console.log(err);
        });

        xlsx.generate(out);

      });
    });
    event.retrunValue = "OK";
  });

  

  function makeBill(writer){

    sheet = xlsx.makeNewSheet();
    sheet.name = writer["name"];
    sheet.columnsWidth = [];
    for(var i=0; i<22; i++)
      sheet.columnsWidth[i] = 1.8;

    sheet.setCell('B2', '御　請　求　書');
    sheet.setCellStyle(1, 1, '20B');

    sheet.setCell('B4', '株式会社 Ｄｏｎｕｔｓ 御中');
    sheet.setCellStyle(3, 1, '18BU');

    // 日付

    sheet.setCell('Q2', '平成26年x月x日');
    sheet.setCellStyle(1, 16, '14BU');
    sheet.mergeCells([1,16],[1,20]);

    // 請求番号

    sheet.setCell('Q1', '請求番号: xxxxx');
    sheet.setCellStyle(0, 16, '11');
    sheet.mergeCells([0,16],[0,19]);

    sheet.setCell('B6', '下記のとおり御請求申し上げます');
    sheet.setCellStyle(5, 1,'12');

    // 住所

    sheet.setCell('O7', '〒167-0043');
    sheet.setCellStyle(6, 14, '12');

    sheet.setCell('O8', '東京都杉並区上荻');
    sheet.setCellStyle(7, 14, '12');

    sheet.setCell('O9', '4－14－8');
    sheet.setCellStyle(8, 14, '12');

    sheet.setCell('O10', writer["name"]);
    sheet.setCellStyle(9, 14, '12');

    // 銀行

    sheet.setCell('B11', '振込先銀行');
    sheet.mergeCells([10,1],[10,3]);
    sheet.setCellStyle(10, 1, '14BU');

    sheet.setCell('B13', '口座番号');
    sheet.mergeCells([12,1],[12,3]);
    sheet.setCellStyle(12, 1, '14BU');

    sheet.setCell('B14', '名義');
    sheet.mergeCells([13,1],[13,3]);
    sheet.setCellStyle(13, 1, '14BU');

    sheet.setCell('E11', '三菱東京UFJ銀行');
    sheet.mergeCells([10,4],[10,8]);
    sheet.setCellStyle(10, 4, '14BU');

    sheet.setCell('E12', '西荻窪駅前支店');
    sheet.mergeCells([11,4],[11,8]);
    sheet.setCellStyle(11, 4, '14BU');

    sheet.setCell('E13', '0998683');
    sheet.mergeCells([12,4],[12,8]);
    sheet.setCellStyle(12, 4, '14BU');

    sheet.setCell('E14', 'オオサワリエ');
    sheet.mergeCells([13,4],[13,8]);
    sheet.setCellStyle(13, 4, '14BU');

    // 金額
    sheet.setCell('B16', '合計金額');
    sheet.mergeCells([15,1],[15,7]);
    sheet.setCellStyle(15, 1, '24C');

    sheet.setCell('I16', '¥5,387');
    sheet.mergeCells([15,8],[15,11]);
    sheet.setCellStyle(15, 8, '24C');

    sheet.setCell('B17', '摘要');
    sheet.mergeCells([16,1],[16,7]);
    sheet.setCellStyle(16, 1, '12C');

    sheet.setCell('I17', '数量');
    sheet.mergeCells([16,8],[16,9]);
    sheet.setCellStyle(16, 8, '12C');

    sheet.setCell('K17', '単価');
    sheet.mergeCells([16,10],[16,12]);
    sheet.setCellStyle(16, 10, '12C');

    sheet.setCell('N17', '金額(税抜)');
    sheet.mergeCells([16,13],[16,17]);
    sheet.setCellStyle(16, 13, '12C');

    sheet.setCell('S17', '備考');
    sheet.mergeCells([16,18],[16,20]);
    sheet.setCellStyle(16, 18, '12C');

    sheet.setCell('C18', '原稿料(x月分)');
    sheet.setCell('I18', 'x');
    sheet.setCell('K18', 'yyyy');
    sheet.setCell('N18', 'zzzz');
    sheet.setCell('S18', '');

    for(var i=18; i<31; i++){
      if(i-17 <= 10){
        sheet.setCell('B'+i, ''+(i-17));
        sheet.setCellStyle(i-1, 1, '12C');
      }
      for(var t=1;t<20; t++){
        sheet.setCellStyle(i-1, t, '12C');
      }
      if(i!=30)
        sheet.mergeCells([i-1,2],[i-1,7]);
      else
        sheet.mergeCells([i-1,1],[i-1,7]);
      sheet.mergeCells([i-1,8],[i-1,9]);
      sheet.mergeCells([i-1,10],[i-1,12]);
      sheet.mergeCells([i-1,13],[i-1,17]);
      sheet.mergeCells([i-1,18],[i-1,20]);
    }
    sheet.setCell('C28', '小計');
    sheet.setCell('C29', '源泉徴収税(10.21%)');
    sheet.setCell('N28', '¥6,000');
    sheet.setCell('N29', '¥-613');
    sheet.setCell('B30', '合計');
    sheet.setCell('N30', '¥5,387');

  }

  // and load the index.html of the app.
  mainWindow.loadUrl('file://' + __dirname + '/index.html');

  // Emitted when the window is closed.
  mainWindow.on('closed', function() {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    mainWindow = null;
  });
});
