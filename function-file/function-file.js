'use strict';

var item;

Office.initialize = function(){
  item = Office.context.mailbox.item;
}

function renderMarkdownQuick(event) {
  hideMessage('error');

  item.body.getTypeAsync(
    function (result) {
      if (result.status == Office.AsyncResultStatus.Failed){
        console.log(result.error.message);
        event.completed();
      }
      else {
        try {
          item.getSelectedDataAsync('text', function (ar) {
            if (isEmptyOrWhiteSpace(ar.value.data)) {
              // Convert entire doc??
              showErrorMessage('error', 'No text selected. Please select markdown to convert and try again.');
              event.completed();
              return;
            }

            convertMarkdown(ar.value.data, result.value, event, item.body.setSelectedDataAsync);
          });
        }
        catch (err) {
          showErrorMessage("error", err.message);
          event.completed();
        }
      }
  });
}

function convertMarkdown(text, format, event, writeFunction) {
  if (text) {
    // Async highlighting with pygmentize-bundled
    marked.setOptions({
      highlight: function (code, lang, callback) {
          try{
              callback(null, code);
          }
          catch(err){
            showErrorMessage("error", err.message);
            event.completed();
            return;
          }
      }
    });

    marked.sanitize = true;
    marked.smartLists = true;

    // Using async version of marked
    marked(text, function (err, content) {
      if (err) {
        showErrorMessage("error", err.message);
        event.completed();
        return;
      }
      
      insertMarkdown(content, format, event, writeFunction);
    });
  }
}


function insertMarkdown(content, format, event, writeFunction) {
    writeFunction(
      content,
      { coercionType: format,
      asyncContext: null },
      function (asyncResult) {
        if (asyncResult.status == 
            Office.AsyncResultStatus.Failed) {
              console.log(asyncResult.error.message);
        }
        else {
        }
        event.completed();
    });
}

function showErrorMessage(key, message) {
    Office.context.mailbox.item.notificationMessages.addAsync(key, {
    type: "errorMessage",
    message : message
  });
}

function showInformationalMessage(key, message) {
    Office.context.mailbox.item.notificationMessages.addAsync(key, {
    type: "informationalMessage",
    message : message,
    icon: 'icon-16',
    persistent: false
  });
}

function hideMessage(key) {
  Office.context.mailbox.item.notificationMessages.removeAsync(key);
}

function isEmptyOrWhiteSpace(str){
    return str === null || str.match(/^\s*$/ ) !== null;
}