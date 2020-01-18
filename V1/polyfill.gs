// https://tc39.github.io/ecma262/#sec-array.prototype.findindex
if (!Array.prototype.findIndex) {
  Object.defineProperty(Array.prototype, 'findIndex', {
    value: function(predicate) {
     // 1. Let O be ? ToObject(this value).
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If IsCallable(predicate) is false, throw a TypeError exception.
      if (typeof predicate !== 'function') {
        throw new TypeError('predicate must be a function');
      }

      // 4. If thisArg was supplied, let T be thisArg; else let T be undefined.
      var thisArg = arguments[1];

      // 5. Let k be 0.
      var k = 0;

      // 6. Repeat, while k < len
      while (k < len) {
        // a. Let Pk be ! ToString(k).
        // b. Let kValue be ? Get(O, Pk).
        // c. Let testResult be ToBoolean(? Call(predicate, T, « kValue, k, O »)).
        // d. If testResult is true, return k.
        var kValue = o[k];
        if (predicate.call(thisArg, kValue, k, o)) {
          return k;
        }
        // e. Increase k by 1.
        k++;
      }

      // 7. Return -1.
      return -1;
    },
    configurable: true,
    writable: true
  });
}


// this loads the es6-promises polyfill to make promise syntax available in Apps Script
// copyright notice - https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
var Promise,
    setTimeout = setTimeout || function (func,ms) {
      Utilities.sleep(ms);
      func();
    };

(function () {  

  // get the polyfill and eval
  if (!Promise) {
    var result = UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/es6-promise/3.2.1/es6-promise.min.js');
    eval (result.getContentText());

    // add done for compatibility with other promise systems
    Promise.prototype.done = Promise.prototype.done || Promise.prototype.then ;

  }
  
}());
