/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

let updatedValue: string;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("update").onclick = update;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = activeSheet.tables.getItemOrNullObject("Sample");
      table.load("isNullObject");
      await context.sync();
      if (table.isNullObject) {
        table = activeSheet.tables.add("a1:j1", true);
        table.name = "Sample";
        table.getHeaderRowRange().values = [
          ["Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8", "Col9", "Col10"],
        ];

        // add validation data to a column
        const column = table.columns.getItem("col3");
        column.getDataBodyRange().dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "a, b, c",
          },
        };

        const rowCount = 20;

        const tableData: string[][] = [];
        for (let i = 0; i < rowCount; i++) {
          tableData[i] = [
            `${i + 1}`,
            `data2 ${i + 1}`,
            "a",
            `data4 ${i + 1}`,
            `data5 ${i + 1}`,
            `data6 ${i + 1}`,
            `I am a very long text that could probably cause the issue of updating the excel sheet for the rows that are not visible in the view port for excel office online. I am a very long text that could probably cause the issue of updating the excel sheet for the rows that are not visible in the view port for excel office online ${
              i + 1
            }`,
            `data8 ${i + 1}`,
            `data9 ${i + 1}`,
            `data10 ${i + 1}`,
          ];
        }

        table.rows.add(null, tableData);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function update() {
  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = activeSheet.tables.getItemOrNullObject("Sample");
      table.load("isNullObject");
      await context.sync();
      if (table.isNullObject) {
        console.log("table is not present. Please click on run first!");
        return;
      }

      const updatedColumnValue = updatedValue === "a" ? "b" : "a";
      updatedValue = updatedColumnValue;
      const updatedColumnValues = new Array(20).fill(updatedColumnValue).map((i) => [i]);
      table.columns.getItem("Col3").getDataBodyRange().values = updatedColumnValues;
    });
  } catch (error) {
    console.error(error);
  }
}
