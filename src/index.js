'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
        var item = Office.context.mailbox.item;
      //prependItemBody(Office.context.mailbox.item);
      //loadItemProps(Office.context.mailbox.item);
    });
  };
  

})();