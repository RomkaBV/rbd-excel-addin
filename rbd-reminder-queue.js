/* RBD reminder queue - local Excel transport preparation. */
(function(){
 "use strict";
 function build(record,user){
  var email=String((user||{})["Електронна пошта"]||"").trim();
  return {
   id:record.id,requestId:record.requestId,recipient:record.recipient,
   email:email,due:record.due,text:record.text,
   state:email?"Очікує підключення Google":"Немає email"
  };
 }
 window.RBD_REMINDER_QUEUE=Object.freeze({build:build,connected:false});
})();