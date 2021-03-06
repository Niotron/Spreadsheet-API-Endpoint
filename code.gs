let SpreadsheetURL = "Your_Spreadsheet_URL_Here";

/*CRUD Request Type
  |
  + GET - |
  |       +- 1. GET Nth ROW
  |       +- 2. GET Nth Column
  |       +- 3. GET Cell Value
  |
  + POST -|
          +- 1. Append :  This function is defined by method and is for adding/appending new row into the sheet
          +- 2. Update : Update cell values of given row.
          +- 3. Delete : Remove/Delete a row.
          +- 4. Read : Read each sheet of your spreadsheet and return JSON array.
          +- 5. Read nth Sheet | Get all rows of sheet as JSON Array
          
  ----------------------------------------------------
  Author : Tanish Raj
  Created On : 03/01/2021
  Updated On : 06/03/2021 | Added web calls to get data directly from spreadsheet
  Version : 1.0.4
  ----------------------------------------------------        
  
*/

function doGet(e){
  var response = {"code":401,"msg":"Failed to process your request. Check Credentials and try again."};
  try{
    if(e.parameter.method == "getRow"){
       response = getRow(e.parameter.sheetName, e.parameter.row);
    }else if(e.parameter.method == "getColumn"){
       response = getColumn(e.parameter.sheetName, e.parameter.column);
    }else if(e.parameter.method == "getCell"){
       response = getCellValue(e.parameter.sheetName, e.parameter.column, e.parameter.row);
    }
  } catch (e){/*Do nothing*/}
  
  return ContentService.createTextOutput(JSON.stringify(response));
}

/*get Process*/
function getRow(sheetName,srow){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied."};
  try{
    var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
    var row = parseInt(srow) + 0;
    if(row <= sheet.getLastRow() && row > 0){
        response = {"code":200,"msg":"Successfully got row", "data": sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]};
    }else{
       response['msg'] = "Invalid row number.";
     }
  }catch (e){/*Do nothing*/}
  
  return response;
}

/*Get Column*/
function getColumn(sheetName,column){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied."};
  try {
    var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
    var columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var column_index = columns.indexOf(column) +1;
    if(column_index>0){
       var body = sheet.getRange(2, column_index, sheet.getLastRow()-1, 1).getValues();
       
       var data = [];
       for(var i=0; i<body.length; i++){
           var value = body[i][0];
           if(Number.isInteger(value)){
              value = value.toLocaleString('fullwide', {useGrouping:false});
           }
           data.push(value);
       }
       
       response = {"code":200,"msg":"Successfully got column values.","data": data};
       
       
    }else{
      response['msg'] = "Invalid column name.";
    }
    
   } catch(e){/*Do nothing*/}
   
   return response;
}

/*Get Cell*/
function getCellValue(sheetName,column,srow) {
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied."};
  try {
    var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
    var columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var row = parseInt(srow) + 0;
    var column_index = columns.indexOf(column) +1;
    if(column_index>0){
     if(row <= sheet.getLastRow() && row > 0){
        var value  = sheet.getRange(row, column_index).getValue();
        if(Number.isInteger(value)){
           value = value.toLocaleString('fullwide', {useGrouping:false});
        }
        response = {"code":200,"msg":"Successfully got cell value.","data": value};
     }else{
       response['msg'] = "Invalid row number.";
     }
    }else{
      response['msg'] = "Invalid column name.";
    }
    
  } catch(e){/*Do nothing*/}
  
  
  return response;
}

/*POST Functions*/
function doPost(e) {
  var response = {"code":401,"msg":"Failed to process your request. Check Credentials and try again."};
  try{
    if(e.parameter.method == "Append"){
       response = CreateRow(e.parameter.sheetName, e.parameter.content);
    }else if(e.parameter.method == "Update"){
       response = UpdateRow(e.parameter.sheetName, e.parameter.content, e.parameter.row);
    }else if(e.parameter.method == "Delete"){
       response = DeleteRow(e.parameter.sheetName, e.parameter.row);
    }else if(e.parameter.method == "Read"){
       response = getSpreadsheet("Successfully got spreadsheet data.");
    }else if(e.parameter.method == "GETSHEET"){
       response = getnSheet(e.parameter.sheetName);
    }
  } catch (e){/*Do nothing*/}
  
  return ContentService.createTextOutput(JSON.stringify(response));
}

/*getnSheet*/
function getnSheet(sheetName){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied."};
  try{
    var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
    response = {"code":200,"msg":"successfully got sheet", "data": sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues()};
  }catch (e){/*Do nothing*/}
  
  return response;
}

function CreateRow(sheetName,sdata){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied Or the JSON is roughly formatted. Check and try again"};
  try {
     var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
     var temp = JSON.parse(sdata);
     for(var i=0;i<temp.length;i++){
         sheet.appendRow(temp[i]);
     }
     response = {"code":200,"msg":"Row added successfully"};
   } catch (e){/*Do nothing*/};
    
   return response;
}


function DeleteRow(sheetName,srow){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied."};
  try{
    var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
    var row = parseInt(srow) + 0;
    if(row <= sheet.getLastRow() && row > 0){
        sheet.deleteRow(row);
        response = {"code":200,"msg":"Row deleted Successfully."};
    }else{
       response['msg'] = "Invalid row number.";
     }
  }catch (e){/*Do nothing*/}
  
  return response;
}

function UpdateRow(sheetName,sdata,srow){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied or the JSON is roughly formatted. Check and try again"};
  try {
     var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sheetName);
     var temp = JSON.parse(sdata)[0];
     var row = parseInt(srow) + 0;
     
     /*Check Row existence*/
     if(row <= sheet.getLastRow() && row > 0){
        var columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var ucolms = Object.keys(temp);
        
        /*Check if the keys given by user are present in sheet*/
        var flag = true;
        for(var i=0; i<ucolms.length;i++){
           if(flag){
              if(columns.indexOf(ucolms[i]) < 0){
                flag = false;
                response['msg'] = "Invalid column given.";
               }
            }
         }
         
         /*Check flag & Store data*/
         if(flag){
            for(var c=0; c<ucolms.length; c++){
               sheet.getRange(row, columns.indexOf(ucolms[c]) +1).setValue(temp[ucolms[c]]);
               response = {"code":200,"msg":"Row updated Successfully."};
            }
         }
        
     }else{
        response['msg'] = "Invalid row number.";
     }
     
     } catch (e){/*Do nothing*/}
     
  return response;
}


function getSpreadsheet(msg){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied.","data":[]};
  var rdata = getSheetList();
  if(rdata['code'] == 200){
     var sheets = rdata['data'];
     var body = sheets.map(r=>{
         let obj = {};
         obj[r] = getSheet(r);
         return obj;
     });
     
     response['code'] = 200;
     response['msg'] = msg;
     response['data'] = body;
     response['sheets'] = rdata['data'];
  }
  
  return response;
}

/*Get Sheet | Type = Internal Function*/
function getSheet(sname){
   var sheet = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheetByName(sname);
   return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
}


/*Get Sheet List | Type = Internal Function*/
function getSheetList(){
  var response = {"code":401,"msg":"Invalid sheet name or the permission to access the sheet is denied.","data":[]};
  try {
    var sheets = SpreadsheetApp.openByUrl(SpreadsheetURL).getSheets();
    var list = [];
    sheets.forEach(function(val){list.push(val.getName())});
    response["data"] = list;
    response['msg'] = "Successfully got sheets.";
    response['code'] = 200;
  } catch (e){/*Do nothing*/}
  
  return response; 
}
