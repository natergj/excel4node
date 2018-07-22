class SimpleLogger {
  constructor(opts) {
    this.logLevel = opts.logLevel || 5;
  }

  debug() {
    if (this.logLevel >= 5) {
      console.debug(...arguments);
    }
  }

  log() {
    if (this.logLevel >= 4) {
      console.log(...arguments);
    }
  }

  inspect() {
    if (this.logLevel >= 4) {
      console.log(...arguments);
    }
  }

  info() {
    if (this.logLevel >= 3) {
      console.info(...arguments);
    }
  }

  warn() {
    if (this.logLevel >= 2) {
      console.warn(...arguments);
    }
  }

  error() {
    if (this.logLevel >= 1) {
      console.error(...arguments);
    }
  }

}

module.exports = SimpleLogger;