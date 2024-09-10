﻿/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

let workTableHeaders;

// #region Any functions that will run exclusively in home.js ----------------------------------------------------------------------------------------

function removeDuplicates(arr) {
  return arr.filter((item,
      index) => arr.indexOf(item) === index);
}

/**
 * Calculates the decimal hours between two timestamps, rounded to the nearest 5 in the decimal hundredths place.
 *
 * @param {string} timestamp1 - The first timestamp in ISO 8601 format.
 * @param {string} timestamp2 - The second timestamp in ISO 8601 format.
 * @returns {number} The decimal hours between the two timestamps, rounded to the nearest 5 in the decimal hundredths place.
 */
function calculateDecimalHours(timestamp1, timestamp2) {
  // Convert timestamps to milliseconds
  const milliseconds1 = new Date(timestamp1).getTime();
  const milliseconds2 = new Date(timestamp2).getTime();

  // Calculate the difference in milliseconds
  const millisecondsDifference = milliseconds2 - milliseconds1;

  // Convert milliseconds to hours (1 hour = 3600000 milliseconds)
  const decimalHours = millisecondsDifference / 3600000;
	const roundedDecimalHours = Math.round(decimalHours * 2) / 2;

  return roundedDecimalHours;
}

/**
 * Converts the provided column index number to it's appropriate column letter
 * @param {Number} index The column index that you are trying to convert to a column letter
 * @returns String
 */
function printToLetter(index) {
  // index -= 1;
  const alphabet = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
  let letter = "";

  while (index >= 0) {
    letter = alphabet[index % 26] + letter;
    index = Math.floor(index / 26) - 1;
  }

  return letter;

};
/**
 * Picks a random value from an array.
 * @param {*} array 
 * @returns {*} A random value.
 */
function randomItem(array) {
  const randomIndex = Math.floor(Math.random() * array.length);
  return array[randomIndex];
}

/**
 * Turns a string to Title Case
 * @param {string} str 
 * @returns {string} A Title Cased String
 */
function toTitleCase(str) {
  return str.replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

// #endregion ----------------------------------------------------------------------------------------------------------------------------------------

import globalVar from "./globalVars.js";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // Check for if we are on a breakout sheet
    document.getElementById("loading").style.display = "none";
    checkSheet();
  }

});

// #region Excel.run Functions -----------------------------------------------------------------------------------------------------------------------

/**
 * Updates a row in the Master row, and then updates any other worksheet in the document.
 * @param {string} ujid The UJID of the row. This unique code will help identify which row we're going to be updating.
 * @param {string} tableName The name of the current table we're on
 * @param {string} operator The name of the pressman.
 * @param {boolean} start If we're logging start or end time. This is true by default.
 */
async function updateRow(ujid, tableName,operator=null, start = true){

  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let ujidIndex = workTableHeaders.indexOf("UJID");
    let startIndex = workTableHeaders.indexOf("Start");
    let stopIndex = workTableHeaders.indexOf("Stop");
    let elapsedIndex = workTableHeaders.indexOf("Elapsed");
    let opIndex = workTableHeaders.indexOf("Operator");

    // In case the tableName argument is Apparel, Digital, Shipping etc...
    let tablesToUpdate = removeDuplicates([tableName,  "Master", "MISSING", "IGNORE", "PRINTED", "DIGITAL", "APPAREL", "Shipping"]);

    for (const table of tablesToUpdate) {
      
      try {
        let currentTable = context.workbook.tables.getItem(globalVar.tableAndSheetNames[table]);
        let dataBodyRange = currentTable.getDataBodyRange();

        // Load values before iterating
        await dataBodyRange.load("values");
        await context.sync();

        let newData = dataBodyRange.values.map((row) => {
          if (row[ujidIndex] == ujid) {
              if (start){
                  // Populate the Start column.
                  row[startIndex] = new Date();
                  row[opIndex] = operator;

                } else {
                  // Populate the End column and the elapsed.
                  row[stopIndex] = new Date();
                  row[elapsedIndex] = calculateDecimalHours(row[startIndex], row[stopIndex]);
                }
          }
          return row;
          // End Map function
        });
        
        dataBodyRange.values = await newData;
        await context.sync(); // Error: Missing, Invalid or incorrect format
        // console.log(`Wrote to ${table}`)
    } catch(e){
          // The table we were looping probably didn't have this row...? idk
      console.log("-".repeat(10), "\n")
      console.log(`Hit a snag in ${table}.`)
      console.error(e)
      console.log("-".repeat(10))
    }

    }
  });
}

/**
 * Checks if Master sheet is present.
 */
async function checkSheet() {
  await Excel.run(async (context) => {
    try {
      const masterSheet = context.workbook.worksheets.getItem("Master").load("name");
      await context.sync();
      // console.log(`Found worksheet name ${masterSheet.name}.`);
      loadJobs();

    } catch (e) {
      document.getElementById("noSheet").style.display = "block";
    }
  });
}

/**
 * Upon succesful locating of the Master worksheet, this function will build cards based on what's needed.
 */
async function loadJobs() {
  Excel.run(async (context) => {
    document.getElementById("jobsPage").style.display = "block";
    
    const workbook = context.workbook;

    const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load("id")
    await context.sync();

    const activeWorksheetId = activeWorksheet.id; // Get the ID of the active worksheet
    await handleWorksheetActivation({ worksheetId: activeWorksheetId });

    workbook.worksheets.onActivated.add(async (eventArgs) => {
      await handleWorksheetActivation(eventArgs);
    });
  }).catch(function (error) {
    console.log("Error:", error);
  });
}

/**
 * Load the jobs when the document opens.
 * @param {*} eventArgs 
 */
async function handleWorksheetActivation(eventArgs) {
  Excel.run(async (context) => {

    $("#jobsPage").empty();
    $("#jobCount, #unauthorized, #completedcollapse").hide();
    const workbook = context.workbook;
    const worksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
    worksheet.load('name');
    await context.sync();

    if (globalVar.ignoreSheets.includes(worksheet.name)) return $("#unauthorized").show();
    // console.log(`Opened sheet: ${worksheet.name}`)

    const workTable = workbook.tables.getItem(globalVar.tableAndSheetNames[worksheet.name]);
    const workTableRange = workTable.getDataBodyRange().load("values"); // Output: array of rows.
    workTableHeaders = workTable.getHeaderRowRange().load("values");
  
    const pressTable = workbook.tables.getItem("Pressmen");
    const pressMenVals = pressTable.getDataBodyRange().load("values");

    await context.sync();
    workTableHeaders = workTableHeaders.values[0];


    // Grab Pressmen
    let pressMen = [];
    pressMenVals.values.forEach((guy)=>{
        pressMen.push(guy[0]);
    });

    // Changes the big header to whatever the line is.
    $("#type").text(toTitleCase(worksheet.name));

    // #region Fill in Cards -------------------------------------------------------------------------------------------------------------------------
    const workHeaders = workTableHeaders;
    
    let jobNumber = 0;
    let completedNumber = 0;
    workTableRange.values.forEach((row) => {
        // Check if the row is empty
        const isEmptyRow = row.every((cellValue) => cellValue === "");

        if (!isEmptyRow) {
            const forms = row[workHeaders.indexOf("Forms")];
            const company = row[workHeaders.indexOf("Company")];
            const total = row[workHeaders.indexOf("Total")];
            const product = row[workHeaders.indexOf("Product")];
            const UJIDArr = (row[workHeaders.indexOf("UJID")]).split("-");

            /* 
            *  This code checks if a job is complete based on the presence of an "End" time in the data.
            *  - If the "End" column is empty (""), the job is considered incomplete and marked with `isComplete = false`.
            *  - If there's a value in the "End" column, the job is considered complete (`isComplete = true`) and we can skip further processing.
            */
            const isComplete = row[workHeaders.indexOf("Stop")] == "" ? false : true; 

            if (!isComplete){
              jobNumber++;
              // List this job. It needs done.
              $("#jobsPage").append(
              `<div class="card ms-bgColor-neutralLight">
                  <h3>Form: <span class="form">${forms}</form></h3>
                  <p><small class="product">${product}</small></p>
                  <p>Client: <span class="company">${company}</span></p>  
                    <details>
                    <summary>More details</summary>
                    <p>Total Qty: <span class="quantity">${total.toLocaleString()}</span></p>
                    <p>View Artwork: <span class="link"><a href="https://employee.themailshark.net/addorderlines2.aspx?c=${UJIDArr[1]}&o=${UJIDArr[2]}">[Link]</a></span></p>
                    <p class="ujidSplash">UJID: <span class="UJID">${UJIDArr.join("-")}|${worksheet.name}</span></p>
                    <p><label>Pressman:</label> <select class="op"></select></p><br>
                    <button class="timerButton ms-Button ms-Button--primary"><span class="ms-Button-label">Start Job</span></button>
                  </details> 
                </div>`);
            } else {
              completedNumber++;
              const operator = row[workHeaders.indexOf("Operator")];
              $("#completedJobs").append(
                `<div class="card ms-bgColor-neutralLight">
                  <h3>Form: ${forms}</h3>
                  <p><small>${product}</small></p>
                  <p>Client: ${company}</p>  
                    <details>
                    <summary>More details</summary>
                    <p>Total Qty: ${total.toLocaleString()}</p>
                    <p>View Artwork: <a href="https://employee.themailshark.net/addorderlines2.aspx?c=${UJIDArr[1]}&o=${UJIDArr[2]}">[Link]</a></p>
                    <p>Pressman: ${operator}</p>
                  </details> 
                </div>`);
            }
            
        }
      // End Work Table Range For-Each
    });

    // Add pressmen to all the selects.

    $(".op").each(function() {
      const currentSelect = $(this);
      pressMen.forEach((guy)=>{
        currentSelect.append(`<option value="${guy}">${guy}</option>`)
      });
    })

    // Show the job number at the top.
    $("#jobCount").html(`<h4>There ${jobNumber == 1 ? "is": "are"} <strong>${jobNumber}</strong> job${jobNumber == 1 ? "" : "s"} on this line.</h4>`).show();
    $("#completedcollapse").show();
    $("#completedCount").text(completedNumber);

    // #endregion ------------------------------------------------------------------------------------------------------------------------------------

    // #region Click the Button ----------------------------------------------------------------------------------------------------------------------
    $('.timerButton').off('click'); // Remove any existing event listeners\

    $('.timerButton').click(function (event) {
      event.stopPropagation(); // Prevent event bubbling or capturing

  
      // Access the card element containing the clicked button
      var card = $(this).closest('.card');
      var ujid = card.find('.UJID');
      var currentPressman = card.find('.op');

      // console.log(currentPressman, "finished!")

      var hiddenArr = ujid.text().split("|");
  
      if ($(this).text().includes("Start Job")) {
          // Log to the sheet, starting the timer.
          updateRow(hiddenArr[0],hiddenArr[1],currentPressman.val());


          currentPressman.prop("disabled", true);
          $(this).html("<span class=\"ms-Button-label\">Stop Job</span>");
      } else if ($(this).text().includes("Stop Job")) {
          $(this).remove();
  
          // Log to the sheet, ending the timer. Calculate the elapsed time.
          updateRow(hiddenArr[0], hiddenArr[1],null, false);
  
          // Create a splash text element
          var splashText = $(`<div class="splash-text">${randomItem(["Good job", "Nice one", "Well done", "Excellent", "Way to go"])}, ${currentPressman.val()}!</div>`);
          card.append(splashText);
  
          // Use a timeout to delay removal
          setTimeout(() => {
              // Add the fade-out class to both the card and the splash text
              card.addClass('fade-out');
              splashText.addClass('fade-out');
  
              setTimeout(() => {
                  card.remove();
                  $("#completedJobs").append(
                    `<div class="card ms-bgColor-neutralLight">
                      <h3>Form: ${card.find('.form').text()}</h3>
                      <p><small>${card.find('.product').text()}</small></p>
                      <p>Client: ${card.find('.company').text()}</p>  
                        <details>
                        <summary>More details</summary>
                        <p>Total Qty: ${card.find('.quantity').text()}</p>
                        <p>View Artwork: ${card.find('.link').html()}</p>
                        <p>Pressman: ${currentPressman.val()}</p>
                      </details> 
                    </div>`);
              }, 500); // Fades away and removes the card.
          }, 1500); // Shows the "Good job" text for 1.5 seconds
      }
  });
    // #endregion ------------------------------------------------------------------------------------------------------------------------------------

    // End Excel.run function.
  });
  // End handleWorksheetActivation function.
}
//#endregion -----------------------------------------------------------------------------------------------------------------------------------------

// #region Other Click Events ------------------------------------------------------------------------------------------------------------------------
$("#credits, #closeCredits").on("click", function(){
  let yearString = "";
  let currentYear = new Date().getFullYear() + 5;
  if (currentYear > 2024){
    // It is after 2024
    yearString=`2024-${currentYear}`;
  } else {
    yearString= currentYear;
  }
  $("#currentYear").text(yearString);
  $("#author, #modal-backdrop").toggle();
});
// #endregion ----------------------------------------------------------------------------------------------------------------------------------------