/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// OfficeExtension.config.extendedErrorLogging = true;

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = reviewData;
    document.getElementById("app-body").style.display = "flex";
  }
});


function reviewData() {
  Excel.run(function(context) {
      var sheet = context.workbook.worksheets.getItem("Sample");
      var range = sheet.getRange("A1:A7392").load('values');
      return context.sync()
        .then(() => {
          var arr = range.values;
          var initialData = []
          var groupedData = [];
          var extractedData = [];
          var finalResult = [];

          //insert the data from excel into an array

          for (var i = 0; i < 7392; i++) {
            initialData.push(arr[i]);
          }
          //split the excel data up into arrays of 24 items each
          while (initialData.length) {
            groupedData.push(initialData.splice(0, 24));
          }

          // pull the 4 needed strings for each record and insert into an array
          for (p = 0; p < groupedData.length; p++) {
            if (groupedData[p][7][0].includes("United States")) {
              var company = groupedData[p][6][0];
            } else if (groupedData[p][7][0].includes("Canada")) {
              var company = groupedData[p][6][0];
            } else {
              var company = groupedData[p][7][0];
            }
            extractedData.push(groupedData[p][0][0], groupedData[p][1][0], company,
              groupedData[p][groupedData[p].length - 2][0]);
            // console.log(p);
          }

          // create 2d array with the 4 needed data points for each record
          while (extractedData.length) {
            finalResult.push(extractedData.splice(0, 4));
          }

          //calculate the value of the cell where the data paste will end
          var dataPlacementBegin = (finalResult.length) + 1;
          var dataPlacement= "e2:h" + dataPlacementBegin;

          //run createRow function to insert the data into the spreadsheet
          createRows(finalResult, dataPlacement);

        });

    })

    .catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));

      }
    });
}

//function that creates the rows where data is entered and places the data in the createRows
function createRows(result, position) {
  Excel.run(function(context) {
      var sheet = context.workbook.worksheets.getItem("Sample");
      var dataPlacement = sheet.getRange(position)
      return context.sync()
        .then(() => {
          dataPlacement.values = result;
        });

    })

    .catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}
