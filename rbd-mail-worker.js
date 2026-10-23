/* Parsing runs outside the Office UI thread; the caller enforces a timeout. */
importScripts("./vendor/mail-parsers.js?v=3.4.9");
self.onmessage=async ({data})=>{try{self.postMessage({info:await self.RBD_MAIL_PARSERS.parse(data.buffer,data.name)});}catch(e){self.postMessage({error:String(e.message||e)});}};
