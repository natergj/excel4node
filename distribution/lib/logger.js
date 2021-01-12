"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var SimpleLogger = function () {
  function SimpleLogger(opts) {
    _classCallCheck(this, SimpleLogger);

    this.logLevel = opts.logLevel || 5;
  }

  _createClass(SimpleLogger, [{
    key: "debug",
    value: function debug() {
      if (this.logLevel >= 5) {
        var _console;

        (_console = console).debug.apply(_console, arguments);
      }
    }
  }, {
    key: "log",
    value: function log() {
      if (this.logLevel >= 4) {
        var _console2;

        (_console2 = console).log.apply(_console2, arguments);
      }
    }
  }, {
    key: "inspect",
    value: function inspect() {
      if (this.logLevel >= 4) {
        var _console3;

        (_console3 = console).log.apply(_console3, arguments);
      }
    }
  }, {
    key: "info",
    value: function info() {
      if (this.logLevel >= 3) {
        var _console4;

        (_console4 = console).info.apply(_console4, arguments);
      }
    }
  }, {
    key: "warn",
    value: function warn() {
      if (this.logLevel >= 2) {
        var _console5;

        (_console5 = console).warn.apply(_console5, arguments);
      }
    }
  }, {
    key: "error",
    value: function error() {
      if (this.logLevel >= 1) {
        var _console6;

        (_console6 = console).error.apply(_console6, arguments);
      }
    }
  }]);

  return SimpleLogger;
}();

module.exports = SimpleLogger;
//# sourceMappingURL=logger.js.map