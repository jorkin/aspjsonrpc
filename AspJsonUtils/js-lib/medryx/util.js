dojo.provide("medryx.util");

(function() {
    var util = medryx.util;
    util.UrlDecode = function(psEncodeString) {
        
        var lsRegExp = /\+/g;
      // Return the decoded string
      return unescape(String(psEncodeString).replace(lsRegExp, " ")); 
    };
    
    util.capfirst = function(value){
        // summary: Capitalizes the first character of the value
        value = "" + value;
        return value.charAt(0).toUpperCase() + value.substring(1);
    };
    
    function LAZY() {
        this.toString = function() {
            return "UNINITIALIZED LAZY PROPERTY";
        }
     };
    
     util.LAZY = new LAZY();
    
})();