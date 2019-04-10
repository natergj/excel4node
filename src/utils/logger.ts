// tslint:disable:no-console

export type LogLevel = 'silent' | 'error' | 'warn' | 'info' | 'log' | 'debug';

const logLevelMap = {
  silent: 0,
  error: 1,
  warn: 2,
  info: 3,
  log: 4,
  debug: 5,
};

export interface ILogger {
  debug: (...additionalParams: any[]) => void;
  log: (msg?: any, ...additionalParams: any[]) => void;
  info: (msg?: any, ...additionalParams: any[]) => void;
  warn: (msg?: any, ...additionalParams: any[]) => void;
  error: (msg?: any, ...additionalParams: any[]) => void;
}

export class SimpleLogger implements ILogger {
  logLevel: LogLevel;

  constructor(logLevel: LogLevel = 'silent') {
    this.logLevel = logLevel;
  }

  debug(...args) {
    const e = new Error();
    const callerLine = e.stack.split(' at ')[2];
    if (logLevelMap[this.logLevel] >= 5) {
      console.debug('[DEBUG]', ...args);
      console.debug(` => at ${callerLine}`);
    }
  }

  log(...args) {
    if (logLevelMap[this.logLevel] >= 4) {
      console.log('[LOG]', ...args);
    }
  }

  info(...args) {
    if (logLevelMap[this.logLevel] >= 3) {
      console.info('[INFO]', ...args);
    }
  }

  warn(...args) {
    if (logLevelMap[this.logLevel] >= 2) {
      console.warn('[WARN]', ...args);
    }
  }

  error(...args) {
    if (logLevelMap[this.logLevel] >= 1) {
      console.error('[ERROR]', ...args);
    }
  }
}
