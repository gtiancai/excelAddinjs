/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var msgDiv = document.getElementById("msg");

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("login").onclick = login;
  }
});

export async function login() {
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
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var ranges = sheet.getRange("A1:B100");
      var nameRange = ranges.getCell(0,0);
      var pwdRange = ranges.getCell(0, 1);
      // nameRange.values = [[userName]];
      // pwdRange.values = [[pwd]];
      // msgDiv.innerText = userName + pwd;

      // Read the range address
      // range.load("address");

      var jsforce = require('jsforce');

      var conn = new jsforce.Connection();
      conn.login(userName, pwd+token, function(err, res) {
        if (err) {
          msgDiv.innerText = err;
          return console.error(err);
        }
        conn.query('SELECT Id, Name FROM Account LIMIT 5', function(err, res) {
          if (err) {
            msgDiv.innerText = err;
            return console.error(err);
          }
          console.log(res);
          // msgDiv.innerText = "query retured.";
          // ranges.getCell(0, 0).values = [["Id"]];
          // ranges.getCell(0, 1).values = [["Name"]];
          // msgDiv.innerText = "result: " + JSON.stringify(res);
          msgDiv.innerText += "\n\r size: " + res.totalSize;
          for (var i = 0; i < res.totalSize; i++) {
            msgDiv.innerText = i;
            ranges.getCell(i, 0).values = [["A"]];
            // ranges.getCell(i, 0).values = [[i]];
            // ranges.getCell(i, 1).values = [[res[i].Name]];
          }

          // await context.sync();
        });
      });


      ranges.getCell(10, 0).values = [["Login is called"]];
      await context.sync();
      console.log(`The range address was ${range.address}.`);


    });
  } catch (error) {
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