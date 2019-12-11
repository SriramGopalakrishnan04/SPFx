/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */
    context.document.body.insertParagraph("Hello test", Word.InsertLocation.end);
    $.ajax({
      url: "https://geton365.onmicrosoft.com/_api/web/lists/getbytitle('Project Requests')/items",
      method: "GET",
      headers: { "Accept": "application/json; odata=verbose" },
      success: function (data) {
        context.document.body.insertParagraph("success", Word.InsertLocation.end);
           if (data.d.results.length > 0 ) {
                //This section can be used to iterate through data and show it on screen
                context.document.body.insertParagraph(data.d.results[0].Title, Word.InsertLocation.end);
           }       
     },
     error: function (data) {
      context.document.body.insertParagraph("Error", Word.InsertLocation.end);
    }
});

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);


    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}