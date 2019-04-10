# Workbook

## Create a New Workbook

<!-- tabs:start -->

#### ** CommonJS **

```javascript
const xl = require('excel4node');
const wb = new Workbook();
```

#### ** ESM **

```javascript
import { Workbook } from 'excel4node';
const wb = new Workbook();
```

#### ** TypeScript **

```typescript
import { Workbook } from 'excel4node';
const wb = new Workbook();
```

<!-- tabs:end -->

### Constructor

```javascript
const wb = new Workbook(options?: IWorkbookOptions): Workbook;
```

### Options

The Workbook constructor accepts a single configuration object.

```javascript
IWorkbookOptions {
  jszip?: Partial<JSZipFileOptions>;
  logger?: ILogger;
  logLevel?: LogLevel;
  defaultFont?: Partial<Font>;
  dateFormat?: string;
  defaultWorkbookView?: Partial<WorkbookView>;
  workbookProperties?: Partial<WorkbookProperties>;
}
```

#### jszip

Options passed to the jszip generate function. Available options can be found in the [JSZip Documentation](https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html).

##### excel4node Defaults

```
{
  type: 'uint8array',
  streamFiles: true,
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  compression: 'DEFLATE',
}
```

#### Logger

The Logger can be useful for troubleshooting issues with excel4node. By default, all log messages are suppressed. A
basic logger is built into the library, but can be overridden. A Logger must be a function with the `error`, `warn`,
`info`, `log` and `debug` methods. See
[logger.js](https://github.com/natergj/excel4node/blob/master/source/lib/logger.js) for a sample logger.

#### LogLevel

One of `error`, `warn`, `info`, `log` and `debug`

#### defaultFont

The default font is the font style applied to the workbook cells when no other style is specified.

[Font Type](types.md#Font)

##### excel4node Default

```
{
  color: 'FF000000',
  name: 'Calibri',
  size: 12,
  family: 'roman',
}
```
