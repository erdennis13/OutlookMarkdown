'use strict';

var item;

Office.initialize = function(){
  item = Office.context.mailbox.item;

}

function renderMarkdownQuick(event) {
  item.body.getTypeAsync(
    function (result) {
      if (result.status == Office.AsyncResultStatus.Failed){
        console.log(result.error.message);
        event.completed();
      }
      else {
        item.body.getAsync(
          result.value,
          {asyncContext:null},
          function callback(text) {
            if (text.value) {
              var bodyHtml = /<body.*?>([\s\S]*)<\/body>/.exec(text.value)[1];
              var bodyText = $(bodyHtml).text();

              marked.sanitize = true;
              marked.smartLists = true;
              var html = marked(bodyText);
              
              item.body.prependAsync(
                html,
                { coercionType: result.value,
                asyncContext: null },
                function (asyncResult) {
                  if (asyncResult.status == 
                      Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        event.completed();
                  }
                  else {
                    event.completed();
                  }
              });
            }
          });
      }
  });
}
