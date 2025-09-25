/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("run1").onclick = run1;
    document.getElementById("run2").onclick = run2;
    document.getElementById("run3").onclick = run3;
    document.getElementById("run4").onclick = run4;
    document.getElementById("run5").onclick = run5;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load("address");

      // Update the fill color.
      range.format.fill.color = "black";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run1() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load("address");

      // Update the fill color.
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run2() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load("address");

      // Update the fill color.
      range.format.fill.color = "red";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run3() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load("address");

      // Update the fill color.
      range.format.fill.color = "green";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run4() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load("address");

      // Update the fill color.
      range.format.fill.color = "blue";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run5() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address and values.
      range.load("address, values");

      await context.sync();
      console.log(`The range address was ${range.address}.`);
      console.log(`The range values were ${JSON.stringify(range.values)}.`);
      // Display the message in the task pane instead of using alert
      const messageElement = document.getElementById("message");
      if (messageElement) {
        messageElement.textContent = `单元格 ${range.address} 的值是 ${JSON.stringify(range.values)}`;
      } else {
        console.log(`单元格 ${range.address} 的值是 ${JSON.stringify(range.values)}`);
      }
    });
  } catch (error) {
    console.error(error);
  }
}