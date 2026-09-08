# SheetNext Recipes

Short patterns for common tasks. Use [the API index](core-api.md) for signatures and [integration notes](integration.md) for setup, edition limits and persistence.

## Initialize

```js
import SheetNext from 'sheetnext';
import 'sheetnext/dist/sheetnext.css';
import zhCN from 'sheetnext/locales/zh-CN';

SheetNext.registerLocale('zh-CN', zhCN);
const container = document.querySelector('.sheetnext-container');
const SN = new SheetNext(container, { locale: 'zh-CN' });
```

Provide a mounted `.sheetnext-container` with explicit width and height. Use the component container ref for multiple instances; share no global workbook variable between them.

## Insert A Template

```js
const sheet = SN.activeSheet;
const template = [
  [{ v: 'Title', mr: 2, b: true, s: 16, h: 36 }, '', ''],
  ['Name', 'Amount', 'Remark'],
  ['Example', 100, '']
];

sheet.insertTemplate(template, 'A1', { border: true, align: 'center', width: 120 });
```

## Read And Write Cells

```js
const sheet = SN.activeSheet;
sheet.getCell('A1').editVal = 'Name';
sheet.getCell('B1').editVal = 'Amount';
sheet.getCell('B2').editVal = 123;
sheet.getCell('B2').numFmt = '#,##0.00';
sheet.getCell('B3').editVal = '=SUM(B2:B2)';
console.log(sheet.getCell('B3').calcVal);
```

## Listen To Events

```js
const handleSelection = (event) => console.log(event.data.newCell);
SN.Event.on('afterSelectionChange', handleSelection);
```

When detaching this integration, call `SN.Event.off('afterSelectionChange', handleSelection)`.

## Filter Text

```js
SN.activeSheet.AutoFilter.setRange('A1:D100');
SN.activeSheet.AutoFilter.setColumnFilter(1, {
  type: 'custom',
  filters: [{ operator: 'contains', value: 'East' }],
  and: true
});
```

## JSON Import And Export

```js
const data = await SN.IO.getData();
if (!SN.IO.setData(data)) throw new Error('Workbook restore failed; recovery may be required');
```

This demonstrates a round trip only. Follow [JSON replacement precautions](json-format.md#replacement-and-recovery) before replacing user data.

## File Export

Requires the commercial distribution and its applicable license. Open Edition supports JSON `getData`/`setData`, but disables file import/export.

```js
await SN.IO.export('XLSX');
await SN.IO.export('CSV');
```
