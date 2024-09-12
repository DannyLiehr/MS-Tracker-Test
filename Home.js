/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import globalVar from "./globalVars.js";
import {checkSheet} from "./funcs.js";


Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // Check for if we are on a breakout sheet
    document.getElementById("loading").style.display = "none";
    checkSheet();
  }
});


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