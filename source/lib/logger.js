const util = require('util');
const colors = require('colors');

const curLogLevel = 'DEBUG';

/**
 * Enum for log levels.
 * @enum {number}
 */
const LogLevel = {
    SUPPRESS: 0,
    ERROR: 1,
    WARN: 2,
    INFO: 3,
    DEBUG: 4
};

/**
 * Enum for log colors.
 * @enum {string}
 */
const LogColor = {
    SUPPRESS: 'white',
    ERROR: 'red',
    WARN: 'magenta',
    INFO: 'cyan',
    DEBUG: 'yellow'
};

/*global Levels*/
let Logger = {
    LogLevel() {
        return Levels;
    }
};

for (let level of Object.keys(LogLevel)) {
    if (LogLevel[ level ] > 0) {
        Logger[ level.toLowerCase() ] = function () {
            // Write log only if logging level in the config is above the
            // threshold for the current log level.
            if (LogLevel[ curLogLevel ] >= LogLevel[ level ]) {
                let logDate = new Date().toLocaleString()
                    .replace(/T/, ' ')
                    .replace(/\..+/, '');
                logDate = `[${logDate}]`;
                let logMessage = util.format.apply(util.format, arguments);

                let errStack = (new Error()).stack;
                let stacklist = errStack.split('\n')[2].split('at ')[1];
                let regEx = /\(([^)]+)\)/;
                let parsedStack = regEx.exec(stacklist) ? regEx.exec(stacklist)[1] : stacklist;
                let stackParts = parsedStack.split(':');
                let file = stackParts[0].split('/')[stackParts[0].split('/').length - 1];
                let line = stackParts[1];

                let logTag = colors.blue('[') +
                    colors[ LogColor[ level ] ].bold(level) +
                    colors.blue(']');
                let log = '';
                if (level !== 'INFO') {
                    log = `${logTag}${logDate}[${file}:${line}] ${logMessage}`;
                } else {
                    log = `${logTag}${logDate} ${logMessage}`;
                }
                console.log(log);
            }
        };
    }
}


Logger.inspects = function () {
    if (LogLevel[ curLogLevel ] === LogLevel.DEBUG) {
        let errStack = (new Error()).stack;
        let stacklist = errStack.split('\n')[ 2 ].split('at ')[ 1 ];
        let regEx = /\(([^)]+)\)/;
        let parsedStack = regEx.exec(stacklist) ? regEx.exec(stacklist)[ 1 ] : stacklist;
        let stackParts = parsedStack.split(':');
        let file = stackParts[ 0 ];
        let line = stackParts[ 1 ];
        let tag = colors.blue('[') + colors.green('INSPECTS') + colors.blue(']');
        let logMessage = `${tag}[${file}:${line}]`;

        for (let i = 0; i < arguments.length; i += 1) {
            logMessage += `\n\n ${colors.blue(i.toString())} \n ${util.inspect(arguments[ i ], {
                colors: true,
                depth: null
            })}`;
        }

        console.log(logMessage);
    }

};

Logger.inspect = function () {
    if (LogLevel[ curLogLevel ] === LogLevel.DEBUG) {
        let errStack = (new Error()).stack;
        let stacklist = errStack.split('\n')[ 2 ].split('at ')[ 1 ];
        let regEx = /\(([^)]+)\)/;
        let parsedStack = regEx.exec(stacklist) ? regEx.exec(stacklist)[ 1 ] : stacklist;
        let stackParts = parsedStack.split(':');
        let file = stackParts[ 0 ];
        let line = stackParts[ 1 ];
        let tag = colors.blue('[') + colors.green('INSPECT') + colors.blue(']');
        let logMessage = `${tag}[${file}:${line}]`;

        logMessage += `\n\n ${util.inspect(arguments[ 0 ], {
            colors: true,
            depth: arguments[ 1 ]
        })}`;

        console.log(logMessage);
    }

};

module.exports = Logger;