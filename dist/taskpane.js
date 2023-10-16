/*! For license information please see taskpane.js.LICENSE.txt */
!function(){"use strict";var t,e,r,n={27091:function(t){t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},60806:function(t,e,r){t.exports=r.p+"abdd506fdc0960d5d6dc.css"}},o={};function a(t){var e=o[t];if(void 0!==e)return e.exports;var r=o[t]={exports:{}};return n[t](r,r.exports,a),r.exports}a.m=n,a.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return a.d(e,{a:e}),e},a.d=function(t,e){for(var r in e)a.o(e,r)&&!a.o(t,r)&&Object.defineProperty(t,r,{enumerable:!0,get:e[r]})},a.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),a.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;a.g.importScripts&&(t=a.g.location+"");var e=a.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");if(r.length)for(var n=r.length-1;n>-1&&!t;)t=r[n--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),a.p=t}(),a.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){e=function(){return n};var r,n={},o=Object.prototype,a=o.hasOwnProperty,i=Object.defineProperty||function(t,e,r){t[e]=r.value},c="function"==typeof Symbol?Symbol:{},u=c.iterator||"@@iterator",s=c.asyncIterator||"@@asyncIterator",l=c.toStringTag||"@@toStringTag";function f(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(r){f=function(t,e,r){return t[e]=r}}function p(t,e,r,n){var o=e&&e.prototype instanceof b?e:b,a=Object.create(o.prototype),c=new P(n||[]);return i(a,"_invoke",{value:L(t,r,c)}),a}function h(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}n.wrap=p;var d="suspendedStart",y="suspendedYield",v="executing",m="completed",g={};function b(){}function w(){}function x(){}var k={};f(k,u,(function(){return this}));var E=Object.getPrototypeOf,S=E&&E(E(C([])));S&&S!==o&&a.call(S,u)&&(k=S);var O=x.prototype=b.prototype=Object.create(k);function I(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function j(e,r){function n(o,i,c,u){var s=h(e[o],e,i);if("throw"!==s.type){var l=s.arg,f=l.value;return f&&"object"==t(f)&&a.call(f,"__await")?r.resolve(f.__await).then((function(t){n("next",t,c,u)}),(function(t){n("throw",t,c,u)})):r.resolve(f).then((function(t){l.value=t,c(l)}),(function(t){return n("throw",t,c,u)}))}u(s.arg)}var o;i(this,"_invoke",{value:function(t,e){function a(){return new r((function(r,o){n(t,e,r,o)}))}return o=o?o.then(a,a):a()}})}function L(t,e,n){var o=d;return function(a,i){if(o===v)throw new Error("Generator is already running");if(o===m){if("throw"===a)throw i;return{value:r,done:!0}}for(n.method=a,n.arg=i;;){var c=n.delegate;if(c){var u=A(c,n);if(u){if(u===g)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===d)throw o=m,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=v;var s=h(t,e,n);if("normal"===s.type){if(o=n.done?m:y,s.arg===g)continue;return{value:s.arg,done:n.done}}"throw"===s.type&&(o=m,n.method="throw",n.arg=s.arg)}}}function A(t,e){var n=e.method,o=t.iterator[n];if(o===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=r,A(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),g;var a=h(o,t.iterator,e.arg);if("throw"===a.type)return e.method="throw",e.arg=a.arg,e.delegate=null,g;var i=a.arg;return i?i.done?(e[t.resultName]=i.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=r),e.delegate=null,g):i:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,g)}function T(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function _(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function P(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(T,this),this.reset(!0)}function C(e){if(e||""===e){var n=e[u];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,i=function t(){for(;++o<e.length;)if(a.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=r,t.done=!0,t};return i.next=i}}throw new TypeError(t(e)+" is not iterable")}return w.prototype=x,i(O,"constructor",{value:x,configurable:!0}),i(x,"constructor",{value:w,configurable:!0}),w.displayName=f(x,l,"GeneratorFunction"),n.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===w||"GeneratorFunction"===(e.displayName||e.name))},n.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,f(t,l,"GeneratorFunction")),t.prototype=Object.create(O),t},n.awrap=function(t){return{__await:t}},I(j.prototype),f(j.prototype,s,(function(){return this})),n.AsyncIterator=j,n.async=function(t,e,r,o,a){void 0===a&&(a=Promise);var i=new j(p(t,e,r,o),a);return n.isGeneratorFunction(e)?i:i.next().then((function(t){return t.done?t.value:i.next()}))},I(O),f(O,l,"Generator"),f(O,u,(function(){return this})),f(O,"toString",(function(){return"[object Generator]"})),n.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},n.values=C,P.prototype={constructor:P,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=r,this.done=!1,this.delegate=null,this.method="next",this.arg=r,this.tryEntries.forEach(_),!t)for(var e in this)"t"===e.charAt(0)&&a.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=r)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,o){return c.type="throw",c.arg=t,e.next=n,o&&(e.method="next",e.arg=r),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var i=this.tryEntries[o],c=i.completion;if("root"===i.tryLoc)return n("end");if(i.tryLoc<=this.prev){var u=a.call(i,"catchLoc"),s=a.call(i,"finallyLoc");if(u&&s){if(this.prev<i.catchLoc)return n(i.catchLoc,!0);if(this.prev<i.finallyLoc)return n(i.finallyLoc)}else if(u){if(this.prev<i.catchLoc)return n(i.catchLoc,!0)}else{if(!s)throw new Error("try statement without catch or finally");if(this.prev<i.finallyLoc)return n(i.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&a.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var o=n;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var i=o?o.completion:{};return i.type=t,i.arg=e,o?(this.method="next",this.next=o.finallyLoc,g):this.complete(i)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),_(r),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;_(r)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:C(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=r),g}},n}function r(t,e,r,n,o,a,i){try{var c=t[a](i),u=c.value}catch(t){return void r(t)}c.done?e(u):Promise.resolve(u).then(n,o)}function n(t){return function(){var e=this,n=arguments;return new Promise((function(o,a){var i=t.apply(e,n);function c(t){r(i,o,a,c,u,"next",t)}function u(t){r(i,o,a,c,u,"throw",t)}c(void 0)}))}}function o(t,e){return function(t){if(Array.isArray(t))return t}(t)||function(t,e){var r=null==t?null:"undefined"!=typeof Symbol&&t[Symbol.iterator]||t["@@iterator"];if(null!=r){var n,o,a,i,c=[],u=!0,s=!1;try{if(a=(r=r.call(t)).next,0===e){if(Object(r)!==r)return;u=!1}else for(;!(u=(n=a.call(r)).done)&&(c.push(n.value),c.length!==e);u=!0);}catch(t){s=!0,o=t}finally{try{if(!u&&null!=r.return&&(i=r.return(),Object(i)!==i))return}finally{if(s)throw o}}return c}}(t,e)||a(t,e)||function(){throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}()}function a(t,e){if(t){if("string"==typeof t)return i(t,e);var r=Object.prototype.toString.call(t).slice(8,-1);return"Object"===r&&t.constructor&&(r=t.constructor.name),"Map"===r||"Set"===r?Array.from(t):"Arguments"===r||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(r)?i(t,e):void 0}}function i(t,e){(null==e||e>t.length)&&(e=t.length);for(var r=0,n=new Array(e);r<e;r++)n[r]=t[r];return n}function c(t,e){if(!(t instanceof e))throw new TypeError("Cannot call a class as a function")}function u(e,r){for(var n=0;n<r.length;n++){var o=r[n];o.enumerable=o.enumerable||!1,o.configurable=!0,"value"in o&&(o.writable=!0),Object.defineProperty(e,(void 0,a=function(e,r){if("object"!==t(e)||null===e)return e;var n=e[Symbol.toPrimitive];if(void 0!==n){var o=n.call(e,"string");if("object"!==t(o))return o;throw new TypeError("@@toPrimitive must return a primitive value.")}return String(e)}(o.key),"symbol"===t(a)?a:String(a)),o)}var a}function s(t,e,r){return e&&u(t.prototype,e),r&&u(t,r),Object.defineProperty(t,"prototype",{writable:!1}),t}var l={offer:{cc:"group@ship-around.com",intro:"Dear {name},<br>",body:"Please find attached:<br>",attachments:"{attachments}",note:"If you accept our offer, please note that the last page of our quotation is the proforma invoice.<br>",closing:"Looking forward to your order confirmation."},acknowledge:{cc:"group@ship-around.com",intro:"Dear {name},<br>",body:"Thank you for reaching out to us.<br><br>We have logged your inquiry with reference SALE{lead}.<br>",note:"Please include the above reference in any future correspondence.<br>",closing:"We appreciate your interest and will get back to you shortly.<br><br>If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing your request."},follow_up_1:{cc:"group@ship-around.com",intro:"Dear {name},<br>",body:"Further to our last communication for {quote_items}, we would like to follow up on your interest in pursuing this order.<br>",note:"I have attached our quotation again for your perusal.<br>",closing:"If we can be of any further assistance, please let as know at your earliest convenient."}},f={Q202:"Quotation",DN202:"Delivery Note",PL202:"Packing List",INV202:"Invoice"};Office.onReady((function(t){t.host===Office.HostType.Outlook&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("acknowledge").onclick=d,document.getElementById("prepare-quote-email").onclick=v,document.getElementById("follow-up").onclick=g)}));var p=function(){function t(e){c(this,t),this.item=e}var r,u,p,h;return s(t,[{key:"getEmailContent",value:function(t,e){if(!l[t])throw new Error("No template found for type: ".concat(t));for(var r=l[t],n="",a=0,i=Object.entries(r);a<i.length;a++){var c=o(i[a],2),u=c[0],s=c[1];if("cc"!==u){for(var f=s,p=0,h=Object.entries(e);p<h.length;p++){var d=o(h[p],2),y=d[0],v=d[1];f=f.replace("{".concat(y,"}"),v)}n+=f+"<br>"}}return n}},{key:"addSubject",value:(h=n(e().mark((function t(r){var n,o=this,a=arguments;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=!(a.length>1&&void 0!==a[1])||a[1],t.abrupt("return",new Promise((function(t,e){o.item.subject.getAsync((function(a){var i;a.status===Office.AsyncResultStatus.Failed?e(a.error):(i=n?r+a.value:r,o.item.subject.setAsync(i,(function(r){r.status===Office.AsyncResultStatus.Failed?e(r.error):t()})))}))})));case 2:case"end":return t.stop()}}),t)}))),function(t){return h.apply(this,arguments)})},{key:"addCC",value:(p=n(e().mark((function t(r){var n,o=this,c=arguments;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=c.length>1&&void 0!==c[1]&&c[1],t.abrupt("return",new Promise((function(t,e){o.item.cc.getAsync((function(c){if(c.status===Office.AsyncResultStatus.Failed)e(c.error);else{var u;if(n)u=[r];else{var s=c.value;if(s.includes(r))return void t();u=[].concat(function(t){if(Array.isArray(t))return i(t)}(l=s)||function(t){if("undefined"!=typeof Symbol&&null!=t[Symbol.iterator]||null!=t["@@iterator"])return Array.from(t)}(l)||a(l)||function(){throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}(),[r])}o.item.cc.setAsync(u,(function(r){r.status===Office.AsyncResultStatus.Failed?e(r.error):t()}))}var l}))})));case 2:case"end":return t.stop()}}),t)}))),function(t){return p.apply(this,arguments)})},{key:"addBody",value:(u=n(e().mark((function t(r){var n=this;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,e){n.item.body.prependAsync(r,{coercionType:Office.CoercionType.Html},(function(r){r.status===Office.AsyncResultStatus.Failed?e(r.error):t()}))})));case 1:case"end":return t.stop()}}),t)}))),function(t){return u.apply(this,arguments)})},{key:"displayErrorInTaskpane",value:function(t){var e=document.createElement("div");e.style.color="red",e.textContent=t,document.body.appendChild(e)}},{key:"getDocumentType",value:function(t){for(var e in f)if(t.startsWith(e))return"".concat(f[e]," ").concat(t);return t}},{key:"generateAttachmentTable",value:function(t){var e="";return t&&t.length>0&&(e='<table style="border-collapse: collapse;">',t.forEach((function(t,r){e+='<tr style="padding: 2px; background-color: #f5f5f5;"><td style="border: 1px solid; padding: 2px 4px;">'.concat(r+1,'</td><td style="border: 1px solid gray; padding: 2px 4px;">').concat(t,"</td></tr>")})),e+="</table>"),e}},{key:"listAttachments",value:(r=n(e().mark((function t(){var r=this;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",new Promise((function(t,e){r.item.getAttachmentsAsync((function(n){if(n.status===Office.AsyncResultStatus.Succeeded){var o=n.value.filter((function(t){return t.attachmentType===Office.MailboxEnums.AttachmentType.File&&!t.isInline}));if(o&&o.length>0){var a=o.map((function(t){var e=t.name.split(".").slice(0,-1).join(".");return e=r.getDocumentType(e),r.capitalizeFirstLetter(e)}));t(a)}else console.log("The current message has no file attachments."),t([])}else console.error("Failed to get attachments:",n.error),e(n.error)}))})));case 1:case"end":return t.stop()}}),t)}))),function(){return r.apply(this,arguments)})},{key:"capitalizeFirstLetter",value:function(t){return t.charAt(0).toUpperCase()+t.slice(1)}}]),t}(),h=function(){function t(e,r,n,o){c(this,t),this.modal=document.getElementById(e),this.allInputDivs=Array.from(this.modal.querySelectorAll("div[id$='InputDiv']")),this.inputDivs=r.map((function(t){return document.getElementById(t)})),this.okButton=document.getElementById(n),this.cancelButton=document.getElementById(o),this.setupEventListeners()}return s(t,[{key:"setupEventListeners",value:function(){var t=this;this.okButton.onclick=function(){var e=t.inputDivs.map((function(t){return t.querySelector("input").value}));t.resolve(e),t.clearInputs(),t.hide()},this.cancelButton.onclick=function(){t.reject(new Error("User cancelled the input.")),t.hide()}}},{key:"clearInputs",value:function(){this.inputDivs.forEach((function(t){var e=t.querySelector("input");e&&(e.value="")}))}},{key:"show",value:function(){var t=this;return this.allInputDivs.forEach((function(t){return t.style.display="none"})),this.inputDivs.forEach((function(t){return t.style.display="block"})),new Promise((function(e,r){t.modal.style.display="block",t.resolve=e,t.reject=r,setTimeout((function(){var e=t.inputDivs[0].querySelector("input");e&&e.focus()}),100)}))}},{key:"hide",value:function(){this.modal.style.display="none"}}]),t}();function d(){return y.apply(this,arguments)}function y(){return(y=n(e().mark((function t(){var r,n,a,i,c,u,s,f,d;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,n=Office.context.mailbox.item,r=new p(n),a=new h("inputModal",["leadInputDiv","nameInputDiv"],"modalOk","modalCancel"),t.next=6,a.show();case 6:return i=t.sent,c=o(i,2),u=c[0],s=c[1],t.next=12,r.addSubject("[SALE".concat(u,"] "));case 12:return f=l.acknowledge.cc,t.next=15,r.addCC(f);case 15:return d=r.getEmailContent("acknowledge",{name:s,lead:u}),t.next=18,r.addBody(d);case 18:t.next=23;break;case 20:t.prev=20,t.t0=t.catch(0),r.displayErrorInTaskpane("Error in acknowledgeRFQ: ".concat(t.t0.message));case 23:case"end":return t.stop()}}),t,null,[[0,20]])})))).apply(this,arguments)}function v(){return m.apply(this,arguments)}function m(){return(m=n(e().mark((function t(){var r,n,a,i,c,u,s,f,d,y,v,m;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,n=Office.context.mailbox.item,r=new p(n),a=new h("inputModal",["nameInputDiv"],"modalOk","modalCancel"),t.next=6,a.show();case 6:return i=t.sent,c=o(i,1),u=c[0],t.next=11,r.listAttachments();case 11:if(s=t.sent,f=s.filter((function(t){return t.startsWith("Quotation Q202")})).map((function(t){return t.replace("Quotation ","")})),d="",1===f.length?d="[Quotation ".concat(f[0],"] "):f.length>1&&(d="[Quotations ".concat(f.join(", "),"] ")),!d){t.next=18;break}return t.next=18,r.addSubject(d);case 18:return y=l.offer.cc,t.next=21,r.addCC(y);case 21:return v=r.generateAttachmentTable(s),m=r.getEmailContent("offer",{name:u,attachments:v}),t.next=25,r.addBody(m);case 25:t.next=30;break;case 27:t.prev=27,t.t0=t.catch(0),r.displayErrorInTaskpane("Error in prepareQuoteEmail: ".concat(t.t0.message));case 30:case"end":return t.stop()}}),t,null,[[0,27]])})))).apply(this,arguments)}function g(){return b.apply(this,arguments)}function b(){return(b=n(e().mark((function t(){var r,n,a,i,c,u,s,f,d,y;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,n=Office.context.mailbox.item,r=new p(n),a=new h("inputModal",["leadInputDiv","nameInputDiv","itemsInput"],"modalOk","modalCancel"),t.next=6,a.show();case 6:return i=t.sent,c=o(i,3),u=c[0],s=c[1],f=c[2],t.next=13,r.addSubject("[Follow-up SALE".concat(u,"] "));case 13:return d=l.follow_up_1.cc,t.next=16,r.addCC(d);case 16:return y=r.getEmailContent("follow_up_1",{name:s,quote_items:f}),t.next=19,r.addBody(y);case 19:t.next=24;break;case 21:t.prev=21,t.t0=t.catch(0),r.displayErrorInTaskpane("Error in followUp1: ".concat(t.t0.message));case 24:case"end":return t.stop()}}),t,null,[[0,21]])})))).apply(this,arguments)}}(),t=a(27091),e=a.n(t),r=new URL(a(60806),a.b),e()(r)}();
//# sourceMappingURL=taskpane.js.map