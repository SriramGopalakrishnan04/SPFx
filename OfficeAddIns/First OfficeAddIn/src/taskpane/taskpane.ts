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
  /*$.support.cors = true;
    var url = "https://geton365.sharepoint.com" + "/_api/web/currentuser";
    $.ajax({
      async: true,
        url: url,
        method: "GET",
        headers: {
            Accept: "application/json;odata=verbose"
        },
        xhrFields: { withCredentials: true },
        
        success: function (data) {
            var items = data.d;
           
            const paragraph = context.document.body.insertParagraph(items.LoginName, Word.InsertLocation.end);
            
        },
        error: function (jqxr, errorCode, errorThrown) {
          alert('hu');
            console.log(jqxr.responseText);
            
        }
    }); 
    */

     
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    //paragraph.font.color = "blue";
//alert('hi');
    await context.sync();
  });
}