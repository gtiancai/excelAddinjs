// import { ContextReplacementPlugin } from 'webpack';

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var msgDiv = document.getElementById("msg");
var loginUrl = '';
var conn;
var jsforce = require('jsforce');

Office.onReady(info => {
  msgDiv.innerText = info.host;
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("login").onclick = login;
    document.getElementById("listSObjects").onclick = describeSObjects;
    document.getElementById("retrieveData").onclick = loadOrCreateSheet;
    
    var elements = document.getElementsByName("OrgType");
    for (var i = 0; i < elements.length; i++) {
      elements[i].onclick = changeOrgType;
    }
  }
});

export function changeOrgType(e) {
  msgDiv.innerText = 'changeOrgType clicked';
  var target = (e.target) ? e.target : e.srcElement;
  if (target.value == 'Production') {
    loginUrl = 'https://login.salesforce.com';
  }
  else if (target.value == 'Sandbox') {
    loginUrl = 'https://test.salesforce.com';
  }
  else {
    loginUrl = 'ABC';
  }
  msgDiv.innerText = loginUrl;
}

export function login() {
  var userName = document.getElementById("userName").value;
  var pwd = document.getElementById("password").value;
  var token = document.getElementById("token").value;
  // var userName = $('#userName').val();
  // var pwd = $('password').val();
  // var userName = 'user name';
  console.log('test');
  // console.log('The user name is ${userName}.');
  console.log(pwd);
  try {
    // Excel.run(context => {
      /**
       * Insert your Excel code here
       */
      // var sheet = context.workbook.worksheets.getActiveWorksheet();
      // var ranges = sheet.getRange("A1:B100");

       // ranges.getRow(0).style.shrinkToFit = true;
      // ranges.getRow(1).style.font.load("bold"); // not work
      // var nameRange = ranges.getCell(0,0);
      // var pwdRange = ranges.getCell(0, 1);
      // nameRange.values = [[userName]];
      // pwdRange.values = [[pwd]];
      // msgDiv.innerText = userName + pwd;

      // Read the range address
      // range.load("address");

      conn = new jsforce.Connection({loginUrl: loginUrl});

      conn.login(userName, pwd+token, function(err, res) {
        if (err) {
          msgDiv.innerText = JSON.stringify(err);
          return console.error(err);
        }

        msgDiv.innerText = 'Login successfully.'
        document.getElementById('contentDiv').style.display = "block";
        
        // conn.describeGlobal (function (err, res) {
        //   if (err) {
        //     msgDiv.innerText = err;
        //     return console.error(err);
        //   }
        //   if (!res || !res.sobjects || res.sobjects.length <= 0) {
        //     return;
        //   }
        //   document.getElementById('SObjectsDiv').style.display = "flex";
          
        //   var select = document.getElementById('SObjectList');
        //   select.onchange = function (e) {
        //     var target = e.target ? e.target : e.srcElement;
        //     msgDiv.innerText = target.value;
        //     conn.describe(target.value, function (err, res) {
        //       if (err) {
        //         msgDiv.innerText = err;
        //         return console.error(err);
        //       }
        //       msgDiv.innerText = JSON.stringify(res);

        //       for (var i = 0; i < res.fields.length; i++) {
                
        //       }
        //     });
        //   }
        //   for(var i = 0; i < res.sobjects.length; i++) {
        //     // if (res.sobjects[i].custom) {
        //     if (res.sobjects[i].queryable) {
        //       select.options[select.options.length] = new Option(res.sobjects[i].label, res.sobjects[i].name);
        //     }
        //   }
        // });

//         conn.query('SELECT Id, Name FROM Account LIMIT 5', function(err, res) {
//           if (err) {
//             msgDiv.innerText = err;
//             return console.error(err);
//           }
// /*          
//           var sobj = new SObject(conn, "CustomObject");
//           msgDiv.innerText = JSON.stringify(sobj.describe(function (err, result) {
//             msgDiv.innerText = json.stringify(result);
//           }));
// */
//           console.log(res);
//           // msgDiv.innerText = "query retured.";
//           // ranges.getCell(0, 0).values = [["Id"]];
//           // ranges.getCell(0, 1).values = [["Name"]];
//           // msgDiv.innerText = "result: " + JSON.stringify(res);
//           // msgDiv.innerText += "\n\r size: " + res.totalSize;
//           // var title = ranges.getRange("A1:B1"); // cannot get range with duplicate area with another range??
//           // title.format.fill.color = "#4472C4";
//           // title.format.font.color = "white";

//           ranges.getCell(0,0).format.fill.color = "#4472C4";
//           ranges.getCell(0,0).format.font.color = "white";
//           ranges.getCell(0,0).values = [["Id"]];
//           ranges.getCell(0,1).format.fill.color = "#4472C4";
//           ranges.getCell(0,1).format.font.color = "white";
//           ranges.getCell(0,1).values = [["Name"]];

//           for (var i = 0; i < res.totalSize; i++) {
//             // msgDiv.innerText = i;
//             // ranges.getCell(i, 0).values = [["A"]];
//             ranges.getCell(i + 1, 0).values = [[res.records[i].Id]];
//             // ranges.getCell(i, 1).values = [[res[i].Name]];
//             ranges.getCell(i + 1, 1).values = res.records[i].Name;
//           }

//           context.sync();
//         });
      });


      // ranges.getCell(10, 0).values = [["Login is called"]];
      // context.sync();
      // console.log(`The range address was ${range.address}.`);


    // });
  } catch (error) {
    console.error(error);
  }
}

export function describeSObjects() {
  try {
    conn.describeGlobal (function (err, res) {
      if (err) {
        msgDiv.innerText = err;
        return console.error(err);
      }
      if (!res || !res.sobjects || res.sobjects.length <= 0) {
        return;
      }
      
      var select = document.getElementById('SObjectList');
      // select.onchange = function (e) {
      //   var target = e.target ? e.target : e.srcElement;
      //   msgDiv.innerText = target.value;
      //   conn.describe(target.value, function (err, res) {
      //     if (err) {
      //       msgDiv.innerText = err;
      //       return console.error(err);
      //     }
      //     msgDiv.innerText = JSON.stringify(res);

      //     for (var i = 0; i < res.fields.length; i++) {
            
      //     }
      //   });
      // }
      for(var i = 0; i < res.sobjects.length; i++) {
        // if (res.sobjects[i].custom) {
        if (res.sobjects[i].queryable) {
          select.options[select.options.length] = new Option(res.sobjects[i].label, res.sobjects[i].name);
        }
      }
    });
  }
  catch (err) {
    console.error(err);
  }
}

export async function loadOrCreateSheet() {
  try {
    await Excel.run(async context => {
      var sobjName = document.getElementById('SObjectList').value;

      var sheets = context.workbook.worksheets;
      var isSheetExist = false;
      var sheet;
      sheets.load("items/name");
      // sheets.load("name,position");

      context.sync().then( function () {
        if (sheets.items.length > 1) {
          for (var i in sheets.items) {
            msgDiv.innerText += sheets.items[i].name;
              if (sheets.items[i].name == sobjName) {
                isSheetExist = true;
              }
          }
        }
        
        // var sheet = sheets.getItemOrNullObject(sobjName); // getItem not work
        if (!isSheetExist) {
          sheet = sheets.add(sobjName);
        }
        else {
          sheet = sheets.getItem(sobjName);
        }
        sheet.activate();
        context.sync();
      });
    return context.sync();
    });
  } catch (error) {
    msgDiv.innerText += JSON.stringify(error);
  }
}

export function retrieveData() {
  try {
    var sobjName = document.getElementById('SObjectList').value;
    var soqlStr = 'SELECT Id, Name FROM ' + sobjName + ' LIMIT 10';
sobjName = 'Test';
    Excel.run(context => {
      
      msgDiv.innerText = "+" ;
      var sheets = context.workbook.worksheets;
      var isSheetExist = false;
      var sheet;
        sheets.load("items/name");
        // sheets.load("name,position");
        var stest = sheets[sobjName];
        if (!sheet) {
          msgDiv.innerText = "--Not exist ----";
        }
        else {
          msgDiv.innerText = "--Exist ----";
        }
        context.sync().then( function () {
          
        if (sheets.items.length > 1) {
          for (var i in sheets.items) {
            msgDiv.innerText += sheets.items[i].name;
              if (sheets.items[i].name == sobjName) {
                isSheetExist = true;
              }
          }
        }
         // var sheet = sheets.getItemOrNullObject(sobjName); // getItem not work
        msgDiv.innerText += "," + sheet + ",";
        if (!isSheetExist) {
          msgDiv.innerText += "Not exist";
          // var sheets = context.workbook.worksheets;
          // create a sheet named by SObject API Name
          sheet = sheets.add(sobjName);
          msgDiv.innerText += "created";
        }
        sheet = sheets.getItem(sobjName);
        sheet.activate();
        context.sync();
        //}
        });
        // sheet.activate();
        return;
      conn.query(soqlStr, function(err, res) {
        if (err) {
          msgDiv.innerText = err;
          return console.error(err);
        }
        
        // var sheet = sheets.getItem(sobjName); // getItem not work
        // if (!sheet) {
        //   // var sheets = context.workbook.worksheets;
        //   // create a sheet named by SObject API Name
          sheet = sheets.add(sobjName);
        // }


        // var sheets = context.workbook.worksheets;
        // sheets.load("items/name");
        // for (var item in sheets.items) {
        //   msgDiv.innerText += JSON.stringify(item);
        // }
        // var sheet = sheets.getItem(sobjName);
        // if (sheet.isNullObject) {
        //   msgDiv.innerText = 'not exist ' + sobjName;
        //   var sheet = sheets.add(sobjName);
        // }
        
        sheet.activate();
        sheet.load('name, position');
      context.sync();
        // var table = sheet.tables.getItem(sobjName); // why getItem does not work and will break thie context?
        // if (!table) {
          var table = sheet.tables.add("A1:B1", true);
          table.Name = sobjName;
        // }
        
        table.getHeaderRowRange().values = [["ID", "Name"]];
        
        // var sheet = context.workbook.worksheets.getActiveWorksheet();
        // var rangeStr = "A1:B" + res.totalSize + 1;
        // var range = sheet.getRange(rangeStr);
        // range.load("address");

        // range.getCell(0,0).format.fill.color = "#4472C4";
        // range.getCell(0,0).format.font.color = "white";
        // range.getCell(0,0).values = [["Id"]];
        // range.getCell(0,1).format.fill.color = "#4472C4";
        // range.getCell(0,1).format.font.color = "white";
        // range.getCell(0,1).values = [["Name"]];
        for (var i = 0; i < res.totalSize; i++) {
          // msgDiv.innerText += res.records[i].Id + res.records[i].Name;
          // range.getCell(i, 0).values = [["A"]];
          // range.getCell(i + 1, 0).values = [[res.records[i].Id]];
          // range.getCell(i, 1).values = [[res[i].Name]];
          // range.getCell(i + 1, 1).values = res.records[i].Name;
          // table.rows.add(i + 1, [[res.records[i].Id, res.records[i].Name]]); // NOT work
          table.rows.add(null, [[res.records[i].Id, res.records[i].Name]]);
        }
        
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
          sheet.getUsedRange().format.autofitColumns();
          sheet.getUsedRange().format.autofitRows();
        }
        
        // msgDiv.innerText = "Retrieving is done.";
        context.sync(); // only do sync here does not work too
      });

      context.sync(); // only do sync here does not work
    });
  } catch (error) {
    msgDiv.innerText = error;
    console.error(error);
  }
}

export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      // range.format.fill.color = "yellow";
      range.values=[[5]];

      range.values = [[ 5 ]];
      range.format.autofitColumns();

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}