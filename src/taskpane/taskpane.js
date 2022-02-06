/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
  if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
    console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
  }
  // Assign event handlers and other initialization logic.
  document.getElementById("insert-paragraph").onclick = insertParagraph;
  }
});

function insertParagraph() {
  Word.run(function (context) {

    var docBody = context.document.body;
    docBody.insertParagraph("What a nice paragraph! I sure hope its not just a test paragraph and will not last long...",
                            "Start");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}


