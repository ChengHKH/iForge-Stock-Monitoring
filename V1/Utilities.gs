function checkSize(sheet, range, array){
  if (range.getNumRows() == array.length){
    if (range.getNumColumns() == array[0].length + 1){
      var newRange = range;
    } else{
      var newRange = sheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), array[0].length + 1);
    }
    
  } else if (range.getNumRows() > array.length){
    //delete extra rows
    var deltaRows = range.getNumRows() - array.length;
    sheet.deleteRows(range.getLastRow(), deltaRows);
    
    if (range.getNumColumns() == array[0].length + 1){
      var newRange = sheet.getRange(range.getRow(), range.getColumn(), array.length, range.getNumColumns());
    } else{
      var newRange = sheet.getRange(range.getRow(), range.getColumn(), array.length, array[0].length + 1);
    }
    
  } else if (range.getNumRows() < array.length){
    //add missing rows
    var deltaRows = array.length = range.getNumRows();
    sheet.insertRows(range.getLastRow(), deltaRows)
    
    if (range.getNumColumns() == array[0].length + 1){
      var newRange = sheet.getRange(range.getRow(), range.getColumn(), array.length, range.getNumColumns());
    } else{
      var newRange = sheet.getRange(range.getRow(), range.getColumn(), array.length, array[0].length + 1);
    }
  }
  return newRange;
}


function mergeDuplicate(cells){
  var s = cells.getSheet();
  var counter ={};
  var offset = 0;
  var data = cells.getValues();
  
  data[0].forEach(function(e){counter[e] = (counter[e] || 0) + 1;});
  
  var tracker = "";
  data[0].forEach(function(e){
    if (tracker != e){
      s.getRange(cells.getRow(), cells.getColumn() + offset, 1, counter[e]).merge();
      offset += counter[e];
    }
    tracker = e;
  });
}


function preExist(element, index, array){
  return RegExp(this.replace(/ /g, "")).test(element);
}


function getNamedRange(cells){
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  
  for(i=0;i<namedRanges.length;i++){
    if (cells.getColumn() >= namedRanges[i].getRange().getColumn()
      && cells.getLastColumn() <= namedRanges[i].getRange().getLastColumn()
      && cells.getRow() >= namedRanges[i].getRange().getRow()
      && cells.getLastRow() <= namedRanges[i].getRange().getLastRow()){
      return namedRanges[i];
    }
  }
}


function query(targetRange, parameters){
  var sheet = targetRange.getSheet();
  var name = sheet.getSheetName();
  var spreadsheetID = sheet.getParent().getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetID + '/gviz/tq?headers=1&sheet=' + encodeURIComponent(name) +
    '&range=' + targetRange.getA1Notation() + '&tq=' + encodeURIComponent(parameters) + '&tqx=out:csv';
  
  var token = getToken();
  var response = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + token}}).getContentText();
  
  var data = Utilities.parseCsv(response);
  return data;
}


function getToken(){
  var cache = CacheService.getScriptCache();
  var token = cache.get('OAuth');
  if (token != null){
    return token;
  } else {
    var OAuthToken = ScriptApp.getOAuthToken();
    cache.put('OAuth', OAuthToken, 180);
    return OAuth;
  }
}


function transpose(matrix){
  return Object.keys(matrix[0]).map(function (c) { return matrix.map(function (r) { return r[c]; }); });
}
