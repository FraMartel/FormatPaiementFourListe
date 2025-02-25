/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

async function futFormatPaFourListe(event) {
  // Implement your custom code here. The following code is a simple Excel example.  
  try {
    await Excel.run(async (context) => {
    /** Ajuster la mise en page de la feuille */
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

    // Set print area for selectedSheet to range "A:K"
    selectedSheet.pageLayout.setPrintArea("A:K");
    // Set ExcelScript.PageOrientation.landscape orientation for selectedSheet
    selectedSheet.pageLayout.orientation = Excel.PageOrientation.landscape;
    // Répéter seulement la rangée 5 sur toutes les pages
    selectedSheet.pageLayout.setPrintTitleRows("$5:$5");
    // Set Letter paperSize for selectedSheet
    selectedSheet.pageLayout.paperSize = Excel.PaperType["letter"];
    // Set FitAllColumnsOnOnePage scaling for selectedSheet
    selectedSheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0, scale: null };

    await context.sync();
    });
  } catch (error) {
      // Note: In a production add-in, notify the user through your add-in's UI.
      console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("futFormatPaFourListe", futFormatPaFourListe);
