function editOverview(sheet, cells){
  var cells = sheet.getRangeList(cells.getA1Notation()).getRanges();
  
  cells.forEach(function(cell){
    //get stockSheet
    var namedRange = getNamedRange(cell);
    var namedRangeName = namedRange.getName();
    
    var namedRanges = sheet.getNamedRanges();
    var namedRangesNames = namedRanges
    .map(function(r){return r.getName();})
    .sort();
  
    var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var stockSheetNames = [];
    allSheets.forEach(function(s){
      if (/\w+ Stock/.test(s.getSheetName())){
        stockSheetNames.push(s.getSheetName());
      } else {
        return;
      }
    })
  
    var index = namedRangesNames.findIndex(function(r){if (r == namedRangeName){return true}});
    var stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(stockSheetNames[index]);
  
    //search + replace
    var type = namedRange.getRange().getCell(1, cell.getColumn() - namedRange.getRange().getColumn()).getValue();
    var typeFinder = stockSheet.createTextFinder(type).findAll();
    var typeRange = stockSheet.getRange(typeFinder[0].getA1Notation() + ':' + typeFinder[typeFinder.length - 1].getA1Notation());
  
    var thickness = namedRange.getRange().getCell(2, cell.getColumn() - namedRange.getRange().getColumn()).getValue();
    var thicknessFinder = typeRange.createTextFinder(thickness).findAll();
    var thicknessRange = stockSheet.getRange(thicknessFinder[0].getA1Notation() + ':' + thicknessFinder[thicknessFinder.length - 1].getA1Notation());
  
    var size = namedRange.getRange().getCell(cell.getRow() - namedRange.getRange().getRow(), 2).getValue();
    var sizeRange = thicknessRange.createTextFinder(size).findNext();
    
    var quantityRange = stockSheet.getRange(row, column)
  
    cell.copyTo(something, {contentsOnly:true})
    
  })
  
}


function editStock(sheet){
  var targetRange = sheet.getDataRange();
  var data = query(targetRange, "SELECT A, B, sum(D) GROUP BY A, B PIVOT C");
  var stock = transpose(data);
  var data = query(targetRange, "SELECT A, B, (sum(D)/sum(E)) GROUP BY A, B PIVOT C");
  var level = transpose(data);
  
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var stockSheetNames = [];
  allSheets.forEach(function(s){
    if (/\w+ Stock/.test(s.getSheetName())){
      stockSheetNames.push(s.getSheetName());
    } else {
      return;
    }
  })
  
  var sheetName = sheet.getSheetName();
  var index = stockSheetNames.indexOf(sheetName);
  var name = 'O' + (index + 1) + sheetName.replace(/ /g, "");
  
  var stockRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name);
  var levelRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview Backend').getRange(stockRange.getA1Notation());
  stockRange.offset(0, 1, stock.length, stock[0].length).setValues(stock);
  levelRange.offset(0, 1, level.length, level[0].length).setValues(level);
}


function updateStock(event){
  Logger.log(event.value);
  var sheet = event.source.getActiveSheet();
  var cells = event.range;
  
  var name = sheet.getSheetName();
  
  if (name === 'Overview'){
    editOverview(sheet, cells);
  } else if (/\w+ Stock/.test(name)){
    if (createOverview() !== true){
      editStock(sheet);
    }
  } else {
    return;
  }
}


function createOverview(){
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var stockSheetNames = allSheets.map(function(s){
    if (/\w+ Stock/.test(s.getSheetName())){
      return s.getSheetName();
    }
  });
  
  var Overview = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview');
  var Backend = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview Backend');
  var namedRanges = Overview.getNamedRanges();
  var namedRangesNames = namedRanges
  .map(function(r){return r.getName();})
  .sort();
  
  var createList = [];
  var createListNames = [];
  stockSheetNames.forEach(function(s, index){
    if ('O' + (index + 1) + s.replace(/ /g, "") != namedRangesNames[index]){
      createList.push(s);
      createListNames.push('O' + (index + 1) + s.replace(/ /g, ""));
    }
  })
  
  Logger.log(createList);
  
  if (createList.length > 0){
    createList.forEach(function(s, i){
      var name = createListNames[i];
    
      var targetRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s).getDataRange();
      var data = query(targetRange, "SELECT A, B, sum(D) GROUP BY A, B PIVOT C");
      var stock = transpose(data);
      var data = query(targetRange, "SELECT A, B, (sum(D)/sum(E)) GROUP BY A, B PIVOT C");
      var level = transpose(data);
    
      if (namedRangesNames.some(preExist, s)){
        //if NamedRange already exists
        //slow, need to optimise
        var newIndex = namedRangesNames.findIndex(preExist, name.replace(s.replace(/ /g, ""), ""));
        var namedRange = getNamedRange(SpreadsheetApp.getActiveSpreadsheet().getRangeByName(namedRangesNames[newIndex]));
        var backRange = Backend.getRange(namedRange.getRange().getA1Notation());
        namedRange.getRange().clear();
        backRange.clear();
        var stockRange = checkSize(Overview, namedRange.getRange(), stock);
        var levelRange = checkSize(Backend, backRange, stock);
        namedRange.setRange(stockRange).setName(name);
      
      } else{
        //if NamedRange does not already exist
        if (namedRangesNames[stockSheetNames.indexOf(s) - 1] == undefined){
          var stockRange = Overview.getRange(1, 7, stock.length, stock[0].length + 1);
          var levelRange = Backend.getRange(1, 7, stock.length, stock[0].length + 1);
        
        } else {
          var lastRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(namedRangesNames[i - 1]);
          var stockRange = lastRange.offset(lastRange.getNumRows() + 1, 0, stock.length, stock[0].length + 1);
          var levelRange = Backend.getRange(stockRange.getRow(), stockRange.getColumn(), stockRange.getNumRows(), stockRange.getNumColumns());
        }
      
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(name, stockRange);
        namedRangesNames.push(name)
        namedRangesNames.sort();
        
      }
    
      formatOverview(s, stock, stockRange);
      formatOverview(s, level, levelRange);
      stockLevels(Overview, stockRange.offset(2, 2, stock.length - 2, stock[0].length - 1));
    
    })
    
    return true;
    
  } else{
    return false;
  }
}


function formatOverview(name, stock, range){
  //set values
  range.offset(0, 1, stock.length, stock[0].length).setValues(stock);
  
  //name
  range.offset(0, 0, stock.length, 1)
  .mergeVertically()
  .setValue(name.replace(/ Stock/, ""))
  .setFontWeight('bold')
  .setVerticalAlignment('middle')
  .setTextRotation(90);
  
  if (range.getSheet().getName() == 'Overview'){
    //heading
    range.offset(0, 1, 2, 1)
    .setBackground('#d9d9d9')
    .setBorder(false, false, true, false, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
    //colour/type + thickness
    range.offset(0, 2, 2, stock[0].length - 1)
    .setBackground('#999999')
    .setFontColor('#ffffff')
    .setBorder(false, false, true, false, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    mergeDuplicate(range.offset(0, 2, 1, stock[0].length - 1));
  
    //size
    range.offset(2, 1, stock.length - 2, 1)
    .setBackground('#efefef');
  }
}


function stockLevels(sheet, ranges){
  var rules = sheet.getConditionalFormatRules();
  var cell = ranges.getCell(1,1).getA1Notation();
  var ranges = [ranges];
  
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('= INDIRECT("Overview Backend!" & ADDRESS(ROW(' + cell + '), COLUMN(' + cell + '))) = 0')
  .setBackground('#e67c73')
  .setRanges(ranges)
  .build();
  
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('= INDIRECT("Overview Backend!" & ADDRESS(ROW(' + cell + '), COLUMN(' + cell + '))) > 0')
  .setBackground('#f6b26b')
  .setRanges(ranges)
  .build();
  
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('= INDIRECT("Overview Backend!" & ADDRESS(ROW(' + cell + '), COLUMN(' + cell + '))) > 0.25')
  .setBackground('#ffd666')
  .setRanges(ranges)
  .build();
  
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('= INDIRECT("Overview Backend!" & ADDRESS(ROW(' + cell + '), COLUMN(' + cell + '))) > 0.50')
  .setBackground('#a9d567')
  .setRanges(ranges)
  .build();
  
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('= INDIRECT("Overview Backend!" & ADDRESS(ROW(' + cell + '), COLUMN(' + cell + '))) > 0.75')
  .setBackground('#57bb8a')
  .setRanges(ranges)
  .build();
  
  rules.push(rule5, rule4, rule3, rule2, rule1);
  sheet.setConditionalFormatRules(rules);
}


//function debug(){
//  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
//  
//  console.time('forEachSheetNames');
//  var stockSheetNames = [];
//  allSheets.forEach(function(s){
//    if (/\w+ Stock/.test(s.getSheetName())){
//      stockSheetNames.push(s.getSheetName());
//    } else {
//      return;
//    }
//  })
//  console.timeEnd('forEachSheetNames');
//  Logger.log(stockSheetNames);
//  
//  console.time('filterSheetNames');
//  var stockSheetNames = allSheets
//  .filter(function(s){return /\w+ Stock/.test(s.getSheetName())})
//  .map(function(s){return s.getSheetName()});
//  console.timeEnd('filterSheetNames');
//  Logger.log(stockSheetNames);
//  
//  var Overview = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview');
//  var Backend = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview Backend');
//  
//  console.time('forEachNamedRanges');
//  var namedRanges = Overview.getNamedRanges();
//  var namedRangesNames = [];
//  namedRanges.forEach(function(r){
//    namedRangesNames.push(r.getName());
//  })
//  namedRangesNames.sort();
//  console.timeEnd('forEachNamedRanges');
//  Logger.log(namedRangesNames);
//  
//  console.time('mapNamedRanges');
//  var namedRanges = Overview.getNamedRanges();
//  var namedRangesNames = namedRanges
//  .map(function(r){return r.getName();})
//  .sort();
//  console.timeEnd('mapNamedRanges');
//  Logger.log(namedRangesNames);
//  
//  console.time('forEachCreate');
//  var createList = [];
//  var createListNames = [];
//  stockSheetNames.forEach(function(s, index){
//    if ('O' + (index + 1) + s.replace(/ /g, "") != namedRangesNames[index]){
//      createList.push(s);
//      createListNames.push('O' + (index + 1) + s.replace(/ /g, ""));
//    }
//  })
//  console.timeEnd('forEachCreate');
//   
//}
