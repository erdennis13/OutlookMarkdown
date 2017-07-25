var item;

Office.initialize = function(){
  item = Office.context.mailbox.item;

  $(document).ready(function () {
    setItemBody();
  });
}

function setItemBody() {
  item.body.getTypeAsync(
    function (result) {
      if (result.status == Office.AsyncResultStatus.Failed){
        write (result.error.message);
      }
      else {
        item.body.getAsync(
          result.value,
          {asyncContext:"This is passed to the callback"},
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
                asyncContext: {var3:1, var4:2} },
                function (asyncResult) {
                  if (asyncResult.status == 
                      Office.AsyncResultStatus.Failed) {
                        write(asyncResult.error.message);
                  }
                  else {
                    write(html);
                  }
              });
            }

          });
      }
  });
}

function write(message) {
  document.getElementById('message').innerText += message;
}