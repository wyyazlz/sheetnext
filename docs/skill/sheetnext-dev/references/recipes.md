# SheetNext Recipes

Short patterns for common SheetNext development tasks. Check `core-api.md` for exact signatures before expanding these examples.

## Initialize

```js
import SheetNext from 'sheetnext';
import 'sheetnext.css';

const SN = new SheetNext(document.querySelector('#SNContainer'));
```

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
sheet.getCell('A1').value = 'Name';
sheet.getCell('B1').value = 'Amount';
sheet.getCell('B2').numFmt = '#,##0.00';
```

## Listen To Events

```js
SN.Event.on('afterSelectionChange', (e) => {
  console.log(e.data.newCell);
});
```

## JSON Import And Export

```js
const data = await SN.IO.getData();
SN.IO.setData(data);
```

## File Export

```js
await SN.IO.export('XLSX');
await SN.IO.export('CSV');
```
