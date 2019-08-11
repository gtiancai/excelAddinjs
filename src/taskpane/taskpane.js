/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("Hi").onclick = helloWord;
    document.getElementById("login").onclick = login;
  }
});

export async function login() {
  var userName = document.getElementById("userName").value;
  var pwd = document.getElementById("password").value;
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
      var ranges = sheet.getRange("A2:B2");
      var nameRange = ranges.getCell(0,0);
      var pwdRange = ranges.getCell(0, 1);
      nameRange.values = [[userName]];
      pwdRange.values = [[pwd]];

      // Read the range address
      // range.load("address");

      var jsforce = require('jsforce');

  var conn = new jsforce.Connection();
  conn.login(userName, pwd, function(err, res) {
    if (err) {
      nameRange.values = [[err]];
      return console.error(err);
    }
    conn.query('SELECT Id, Name FROM Account', function(err, res) {
      if (err) {
        nameRange.values = [[err]];
        return console.error(err);
      }
      console.log(res);
      pwdRange.values = [[res]];
    });
  });
  


      await context.sync();
      console.log(`The range address was ${range.address}.`);


    });
  } catch (error) {
    console.error(error);
  }
/*
  var jsforce = require('jsforce');

  var conn = new jsforce.Connection();
  conn.login(userName, pwd, function(err, res) {
    if (err) {
      alert(err);
      return console.error(err);
    }
    conn.query('SELECT Id, Name FROM Account', function(err, res) {
      if (err) {
        alert(err);
        return console.error(err);
      }
      console.log(res);
      alert(res);
    });
  });
  */
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

export async function helloWord() {
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
      range.values = [["Hello World!"]];

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
