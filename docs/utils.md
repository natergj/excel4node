# Utility Functions

The excel4node library provides some utility functions that may be helpful when generating workbooks.

## getExcelAlpha

Translates a column number into the Alpha equivalent used by Excel

`getExcelAlpha = (colNum: number): string`

<!-- tabs:start -->

#### ** CommonJS **

```javascript
const xl = require('excel4node');
const result = xl.getExcelAlpha(5);
// result = "E"
```

#### ** ESM **

```javascript
import { getExcelAlpha } from 'excel4node';
const result = getExcelAlpha(5);
// result = "E"
```

#### ** TypeScript **

```typescript
import { getExcelAlpha } from 'excel4node';
const result = getExcelAlpha(5);
// result = "E"
```

<!-- tabs:end -->

## getExcelCellRef

Translates a column number into the Alpha equivalent used by Excel

`getExcelCellResult = (rowNum: number, colNum: number): string`

<!-- tabs:start -->

#### ** CommonJS **

```javascript
const xl = require('excel4node');
const result = xl.getExcelCellRef(1, 5);
// result = "E1"
```

#### ** ESM **

```javascript
import { getExcelCellRef } from 'excel4node';
const result = getExcelCellRef(1, 5);
// result = "E1"
```

#### ** TypeScript **

```typescript
import { getExcelCellRef } from 'excel4node';
const result = getExcelCellRef(1, 5);
// result = "E1"
```

<!-- tabs:end -->

## getExcelRowCol

Translates a Excel cell representation into row and column numerical equivalents

`getExcelRowCol = (str: string): { row: number, col: number }`

<!-- tabs:start -->

#### ** CommonJS **

```javascript
const xl = require('excel4node');
const result = xl.getExcelRowCol('B3');
// result = { row: 3, col: 2 }
```

#### ** ESM **

```javascript
import { getExcelRowCol } from 'excel4node';
const result = getExcelRowCol('B3');
// result = { row: 3, col: 2 }
```

#### ** TypeScript **

```typescript
import { getExcelRowCol } from 'excel4node';
const result = getExcelRowCol('B3');
// result = { row: 3, col: 2 }
```

<!-- tabs:end -->

## sortCellRefs

Sorter function for an array of Excel reference strings (i.e. ["A1", "BA2", "AAA2"])

`sortCellRefs = (a: string, b: string): string[]`

<!-- tabs:start -->

#### ** CommonJS **

```javascript
const xl = require('excel4node');
const result = ['AAA2', 'A1', 'BA2'].sort(xl.sortCellRefs);
// result = ["A1", "BA2", "AAA2"]
```

#### ** ESM **

```javascript
import { sortCellRefs } from 'excel4node';
const result = ['AAA2', 'A1', 'BA2'].sort(sortCellRefs);
// result = ["A1", "BA2", "AAA2"]
```

#### ** TypeScript **

```typescript
import { sortCellRefs } from 'excel4node';
const result = ['AAA2', 'A1', 'BA2'].sort(sortCellRefs);
// result = ["A1", "BA2", "AAA2"]
```

<!-- tabs:end -->
