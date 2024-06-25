/**
 * Hermes AI Text Sanitizer Outlook Add-in by Eduardo Arana <info@arananet.net>.
 *
 * The server is expected to return processed text in the response from the AI api which then replaces the original selection.
 * In other words, it "sanitizes" the selected text via the remote server.
 * 
 * If no text is selected when the button is clicked, the script does nothing.
 * 
 * If anything fails during this process, the error will be logged to the console, and no changes will be made to the email.
 */

 let item;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, modifySelectedItem);
  }
});

function setItemBody(event) {
  var item = Office.context.mailbox.item;

  // Get the selected content by the user from the email body
  item.getSelectedDataAsync(Office.CoercionType.Html, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      event.completed({allowEvent: false});
      console.log('Action failed. Error: ' + asyncResult.error.message);
      return;
    }

    let currentSelectedContent = asyncResult.value.data;
    //check if the content is select, if not do not do anything.
       if(currentSelectedContent && currentSelectedContent.trim().length > 0){
      // prepare data for ajax call
      //const data = { content_from_outlook: currentSelectedContent };
      fetch('https://41cb-45-250-252-165.ngrok-free.app/rephrase', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ text: currentSelectedContent }), // Encode as JSON
        })
      .then((response) => {
          if (!response.ok) {
              throw new Error('Network response was not ok');
          }
          return response.text();
      })
      .then((fetchedData) => {
         if(fetchedData) {
            let newSelectedContent = fetchedData;
  
          item.body.getTypeAsync((typeAsyncResult) => {
            if (typeAsyncResult.status == Office.AsyncResultStatus.Failed) {
              event.completed({allowEvent: false});
              console.log('Action failed. Error: ' + typeAsyncResult.error.message);
            } else if (typeAsyncResult.value === Office.CoercionType.Html) {
              // Replace the selected text in HTML body.
              item.body.setSelectedDataAsync(
                `${newSelectedContent}`,
                { coercionType: Office.CoercionType.Html },
                (replaceSelectionAsyncResult) => {
                  if(replaceSelectionAsyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Action failed. Error: " + replaceSelectionAsyncResult.error.message);
                  }
                  event.completed({allowEvent: true});
                }
              );
            } else {
              // Replace the selected text in Text body.
              item.body.setSelectedDataAsync(
                newSelectedContent,
                { coercionType: Office.CoercionType.Html },
                (replaceSelectionAsyncResult) => {
                  if(replaceSelectionAsyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Action failed. Error: " + replaceSelectionAsyncResult.error.message);
                  }
                  event.completed({allowEvent: true}); 
                }
              );
            }
          });
        }
      });
      } else {
      event.completed({allowEvent: true}); 
      console.log("No valid content selected.");
    }
    });
}