// 店鋪管理管理の抽出条件
// とんとん3店舗 => 条件: from:(tonton@megasystems.jp) subject:(売上報告)
// ふるふる => 条件: from:(furufuru1513@ask-santyoku.com) subject:(売上情報) 糀屋団四郎様の売上情報(日計)。
// いっぺこーと => 条件: from:(janiigatamirai1515@ask-santyoku.com) subject:(19時の売上情報)

var productsList = [
  '金印味噌1k',
  '金印味噌500',
  '銀印味噌1k',
  '銀印味噌500',
  '三年味噌500',
  '団四郎の甘酒',
  '団四郎の塩糀',
  '手づくり糀',
  'みそ漬け',
];

var productsList2 = [
  '金印味噌1k',
  '金印味噌500',
  '銀印味噌1k',
  '銀印味噌500',
  '三年味噌500',
  '甘酒',
  '塩糀',
  '手作り糀',
  '味噌漬け',
]; 
  
function getMail() {
  var label = '店鋪管理 ';
  var targetSheet = '2018/9迄';
  var checkSheet = '2018確認用';
  var start = 0;
  var max = 500;
  var threads = GmailApp.search('label:' + label + ' is:unread', start, max);
  var messages = GmailApp.getMessagesForThreads(threads);
  var setProductsData = [
    [0, 0, 0, 0, 0, 0, 0, 0, 0],
    [0, 0, 0, 0, 0, 0, 0, 0, 0],
    [0, 0, 0, 0, 0, 0, 0, 0, 0],
    [0, 0, 0, 0, 0, 0, 0, 0, 0],
    [0, 0, 0, 0, 0, 0, 0, 0, 0],
  ];
  var printData = [];
  var checkData = [['', '', '無し', '無し', '無し', '無し', '無し' ]];
  var getDate;

  for(var i = 0; i < messages.length; i++){
    var last = messages[i].length - 1;
    var getFrom = messages[i][last].getFrom();
    var getSubject = messages[i][last].getSubject();
    getDate = messages[i][last].getDate();
    var plainMessage = messages[i][last].getPlainBody();
    var message = alpha(plainMessage).replace(/金印味噌　/g, "金印味噌").replace(/銀印味噌　/g, "銀印味噌").replace(/三年味噌　/g, "三年味噌").replace(/,/g, "").replace(/㎏/g, "k").replace(/g/g, "").replace(/個/g, "").replace(/円/g, "").split(/\s+/);
    if (getFrom.indexOf('tonton@megasystems.jp') > -1) {
      if (getSubject.indexOf('松崎') > -1) {
        setProductsData[0] = extractionData(message, 1);
        checkData[0][2] = plainMessage;
      } else if (getSubject.indexOf('白根') > -1) {
        setProductsData[1] = extractionData(message, 1);
        checkData[0][3] = plainMessage;
      } else if (getSubject.indexOf('新発田') > -1) {
        setProductsData[2] = extractionData(message, 1);
        checkData[0][4] = plainMessage;
      }
    } else if (getFrom.indexOf('janiigatamirai1515@ask-santyoku.com') > -1) {
      setProductsData[3] = extractionData(message, 2);
      checkData[0][5] = plainMessage;
    } else if (getFrom.indexOf('furufuru1513@ask-santyoku.com') > -1) {
      setProductsData[4] = extractionData(message, 2);
      checkData[0][6] = plainMessage;
      Logger.log(plainMessage);
    }
    
    for(var j = 0; j < messages[i].length; j++){
      messages[i][j].markRead(); //メッセージを既読にする
    }
    
    // Logger.log(message);
  }
  
  var targetRow = SpreadsheetApp.getActive().getSheetByName(targetSheet).getLastRow() + 1;
  var checkTargetRow = SpreadsheetApp.getActive().getSheetByName(checkSheet).getLastRow() + 1;
  
  for(var j = 0; j < productsList.length; j++){
    printData[j] = [
      '', '', productsList[j],
      setStockNum(targetRow, 'D', j), '', '', '', salesAverage(targetRow, 'I', j, 7), setProductsData[0][j],
      setStockNum(targetRow, 'J', j), '', '', '', salesAverage(targetRow, 'O', j, 7), setProductsData[1][j],
      setStockNum(targetRow, 'P', j), '', '', '', salesAverage(targetRow, 'U', j, 7), setProductsData[2][j],
      setStockNum(targetRow, 'V', j), '', '', '', salesAverage(targetRow, 'AA', j, 7), setProductsData[3][j],
      setStockNum(targetRow, 'AB', j), '', '', '', salesAverage(targetRow, 'AG', j, 7), setProductsData[4][j],
    ];
  }

  // 2018
  printData[0][0] = getDate;
  printData[0][1] = getDay(getDate);
  for(var i = 0; i < productsList.length; i++){
    SpreadsheetApp.getActive().getSheetByName(targetSheet).getRange(targetRow, 1, j, 33).setValues(printData);
  }
  // 罫線を挿入
  var borderRange = 'A' + String(targetRow - 1 + productsList.length) + ':' + 'AG' + String(targetRow -1 + productsList.length);
  SpreadsheetApp.getActive().getSheetByName(targetSheet).getRange(borderRange).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.DOUBLE);

  // 2018確認用
  checkData[0][0] = getDate;
  checkData[0][1] = getDay(getDate);
  SpreadsheetApp.getActive().getSheetByName(checkSheet).getRange(checkTargetRow, 1, 1, 7).setValues(checkData);
}
  
// 2018の最終行のセルを開く
function onOpen() {
  var targetSheet = '2018/9迄';
  var lastRow = SpreadsheetApp.getActive().getSheetByName(targetSheet).getLastRow();
  SpreadsheetApp.getActive().getSheetByName(targetSheet).setActiveSelection("A" + lastRow); 
}
  
// 2018確認用の最終行のセルを開く
function onOpenCheck() {
  var targetSheet = '2018確認用';
  var lastRow = SpreadsheetApp.getActive().getSheetByName(targetSheet).getLastRow();
  SpreadsheetApp.getActive().getSheetByName(targetSheet).setActiveSelection("A" + lastRow); 
}
  
// 売上平均値抽出
function salesAverage(targetRow, colName, num, days) {
  var days = String(days); 
  var result = '';
  for(var i = 0; i < days; i++){
    var plus = i !== 0 ? '+' : '';
    result += plus + (colName + String(targetRow + num - i*productsList.length));
  }
  return '=' + result;
};
        
// 曜日抽出
function getDay(day) {
  var weekday = [ "日", "月", "火", "水", "木", "金", "土" ] ;
  var now = new Date(day);
  var day = now.getDay();
  return weekday[day];
};

// 在庫抽出
function setStockNum(targetRow, colName, num) {
  var stock = colName + String(targetRow + num - productsList.length);
  var adjustment = getCol(colName, 1) + String(targetRow + num);
  var add = getCol(colName, 2) + String(targetRow + num);
  var sales = getCol(colName, 5) + String(targetRow + num);
  return '=' + stock + '+' + adjustment + '+' + add + '-' + sales;
};

// 列抽出        
function getCol(colName, num) {
  var alphabets = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG'];
  var colNameNum = alphabets.indexOf(colName);
  return alphabets[colNameNum + num];
};

// productData抽出
function extractionData(message, num) {
  var result = [];
  if (num !== 2) {
    for(var i = 0; i < productsList.length; i++) {
      var target = message.indexOf(productsList[i]);
      result.push(target > -1 ? message[target + num] : '0');
    }
  } else {
    for(var i = 0; i < productsList2.length; i++) {
      var target = message.indexOf(productsList2[i]) ;
      result.push(target > -1 ? message[target + num] : '0');
    }
  }
  return result;
};

// 半角英数字変換
function alpha(str) {
  return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 65248);
  });
};