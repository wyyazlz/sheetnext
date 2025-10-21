# SheetNext API 文档

> 详细的类、方法和属性说明

---

## 目录

- [SheetNext Class](#SheetNext) - 工作簿管理：实例化编辑器、添加/删除/切换工作表、导入/导出 Excel、多实例支持等
- [Sheet Class](#Sheet) - 工作表管理：行列管理、合并单元格、区域遍历、排序、批量插入数据、图形对象管理等
- [Cell Class](#Cell) - 单元格管理：读写值/公式、样式设置、数字格式、超链接、数据验证等
- [Row Class](#Row) - 行管理：行高、批量行样式、获取行内单元格、显示/隐藏等
- [Col Class](#Col) - 列管理：列宽、批量列样式、获取列内单元格、显示/隐藏等
- [Drawing Class](#Drawing) - 图形管理：图表、图形、图片的增删改查、设置位置和尺寸、调整图层顺序等
- [Utils Class](#Utils) - 坐标转换：单元格字符串与数字坐标互转、列字母与数字互转、范围格式转换等
- [UndoRedo Class](#UndoRedo) - 历史管理：撤销/重做操作

---

## SheetNext

SheetNext 是主入口类，管理整个电子表格应用。

### 构造函数

```javascript
const SN = new SheetNext(dom: HTMLElement, options?: object)
// 下面所有示例使用此实例进行操作
```

**参数：**
- `dom`: 容器 DOM 元素（必需）
- `options`: 可选配置对象（详见 [README 初始化配置](./README.md#初始化配置)）

### 核心属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `activeSheet` | `Sheet` | 当前激活的工作表（可读写） |
| `sheets` | `Sheet[]` | 所有工作表数组 |
| `sheetNames` | `string[]` | 工作表名称列表（只读） |
| `containerDom` | `HTMLElement` | 编辑器容器元素 |
| `namespace` | `string` | 实例的全局命名空间（如 `SN_0`） |
| `locked` | `boolean` | 是否锁定工作表切换 |

### 工作表管理方法

#### `addSheet(name?: string): Sheet`
添加新工作表，名称可选（自动生成 Sheet1、Sheet2 等）。

```javascript
const newSheet = SN.addSheet("销售数据");
const autoSheet = SN.addSheet(); // 自动命名
```

**规则：** 名称不重复、长度1-31字符、不含特殊符号 `: / \ * ? [ ]`

#### `delSheet(name: string): void`
删除指定工作表（至少保留一个可见工作表）。

```javascript
SN.delSheet("Sheet2");
```

#### `getSheetByName(name: string): Sheet | null`
根据名称获取工作表。

```javascript
const sheet = SN.getSheetByName("Sheet1");
```

#### `getVisibleSheetByIndex(index: number): Sheet`
获取可见工作表（按索引，隐藏工作表不计入）。

```javascript
const firstSheet = SN.getVisibleSheetByIndex(0);
```

### 文件操作

#### `import(file: File): Promise<void>`
导入 Excel 文件（.xlsx 格式）。

```javascript
fileInput.addEventListener('change', (e) => {
  SN.import(e.target.files[0]);
});
```

#### `export(type: string): void`
导出电子表格（目前支持 `"XLSX"`）。

```javascript
SN.export('XLSX');
```

### 其他方法

#### `r(): void`
手动触发画布重新渲染（批量修改后使用）。

```javascript
// 批量修改后刷新
for (let i = 0; i < 100; i++) {
  sheet.getCell(i, 0).editVal = i;
}
SN.r();
```

#### `getLicenseInfo()`
获取 License 信息。

```javascript
const info = SN.getLicenseInfo();
```

### 多实例支持

SheetNext 支持同一页面创建多个独立实例：

```javascript
const editor1 = new SheetNext(document.querySelector('#container1'));
const editor2 = new SheetNext(document.querySelector('#container2'));

console.log(editor1.namespace); // "SN_0"
console.log(editor2.namespace); // "SN_1"
```

---

## Sheet

Sheet 类代表一个工作表。

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `name` | `string` | 工作表名称 |
| `hidden` | `boolean` | 是否隐藏 |
| `merges` | `RangeNum[]` | 合并单元格区域列表 |
| `defaultColWidth` | `number` | 默认列宽（像素） |
| `defaultRowHeight` | `number` | 默认行高（像素） |
| `showGridLines` | `boolean` | 是否显示网格线 |
| `showRowColHeaders` | `boolean` | 是否显示行列标题 |
| `activeCell` | `CellNum` | 当前活动单元格 |
| `activeAreas` | `RangeNum[]` | 当前选中区域 |
| `rows` | `Row[]` | 所有行 |
| `cols` | `Col[]` | 所有列 |
| `xSplit` | `number` | 冻结列数 |
| `ySplit` | `number` | 冻结行数 |
| `drawings` | `Drawing[]` | 图形对象列表（图表、图片等） |

### 基础方法

#### `getRow(r: number): Row`
获取指定行对象。

```javascript
const row = sheet.getRow(0); // 获取第一行
```

#### `getCol(c: number): Col`
获取指定列对象。

```javascript
const col = sheet.getCol(0); // 获取第一列（A列）
```

#### `getCell(r: number, c: number): Cell`
获取指定单元格。

```javascript
const cell = sheet.getCell(0, 0); // 获取 A1 单元格
```

#### `getCellByStr(cellStr: string): Cell`
通过字符串引用获取单元格。

```javascript
const cell = sheet.getCellByStr("A1");
```

#### `rangeStrToNum(rangeStr: string): RangeNum`
将字符串范围转换为数字范围对象。

```javascript
const rangeNum = sheet.rangeStrToNum("A1:C3");
// 返回: {s:{r:0,c:0}, e:{r:2,c:2}}
```

### 区域遍历

#### `eachArea(rangeRef: RangeRef, callback: (r, c, index) => void, reverse?: boolean): void`
遍历指定区域的每个单元格。

```javascript
// 正向遍历
sheet.eachArea("A1:C3", (r, c, index) => {
  const cell = sheet.getCell(r, c);
  console.log(cell.showVal);
});

// 反向遍历（用于删除操作，避免索引混乱）
sheet.eachArea("A:A", (r, c) => {
  if (sheet.getCell(r, c).showVal === "") {
    sheet.delRows(r, 1);
  }
}, true);
```

### 行列操作

#### `showAllHidRows(): void`
显示所有隐藏的行。

```javascript
sheet.showAllHidRows();
```

#### `showAllHidCols(): void`
显示所有隐藏的列。

```javascript
sheet.showAllHidCols();
```

#### `addRows(startR: number, num: number): void`
在指定位置插入行。

```javascript
sheet.addRows(5, 3); // 在第 5 行位置插入 3 行
```

**注意**：多个同时调用时，应反向遍历以避免索引混乱。

#### `addCols(startC: number, num: number): void`
在指定位置插入列。

```javascript
sheet.addCols(2, 2); // 在第 2 列位置插入 2 列
```

#### `delRows(startR: number, num: number): void`
删除指定行。

```javascript
sheet.delRows(5, 3); // 删除从第 5 行开始的 3 行
```

#### `delCols(startC: number, num: number): void`
删除指定列。

```javascript
sheet.delCols(2, 2); // 删除从第 2 列开始的 2 列
```

### 合并单元格

#### `mergeCells(rangeRef: RangeRef): void`
合并指定区域的单元格。

```javascript
sheet.mergeCells("A1:C3");
// 或
sheet.mergeCells({s:{r:0,c:0}, e:{r:2,c:2}});
```

#### `unMergeCells(cellRef: CellRef): void`
取消合并（传入区域内任意单元格引用）。

```javascript
sheet.unMergeCells("A1"); // 取消包含 A1 的合并区域
```

### 排序

#### `rangeSort(sortItems: SortItem[], range?: RangeRef): void`
对指定区域进行排序。

**SortItem 接口：**

```typescript
interface SortItem {
  type: "column" | "row" | "custom";
  order?: "asc" | "desc" | "value"; // type="custom" 时省略
  index: string; // 列/行标签，行从 1 开始，列从 A 开始
  sortData?: any[]; // order="value" 时使用，基于此数据排序
  cb?: (rowsArray: Cell[][], sortIndex: number) => Cell[][]; // type="custom" 时使用
}
```

**示例：按自定义顺序排序**

```javascript
const sheet = SN.activeSheet;
// 除标题外，按 C 列字母顺序排序：A V U T
sheet.rangeSort(
  [{
    type: 'column',
    order: 'value',
    index: 'C',
    sortData: ["A", "V", "U", "T"]
  }],
  {s:{c:0,r:1}, e:{c:sheet.colCount, r:sheet.rowCount}}
);
```

### 批量插入数据

#### `insertTable(data: (ICellConfig | string | number)[][], startCell: CellRef, globalConfig?: object): RangeNum`
在指定位置插入表格数据。

**ICellConfig 接口：**

```typescript
interface ICellConfig {
  v?: string;       // 单元格值
  w?: number;       // 列宽（像素），仅在首行设置
  h?: number;       // 行高（像素），仅在首列设置
  b?: boolean;      // 是否粗体
  s?: number;       // 字体大小
  fg?: string;      // 背景色
  a?: 'l' | 'r' | 'c'; // 对齐方式（left/right/center）
  c?: string;       // 文本颜色
  mr?: number;      // 向右合并单元格数（不包括自身）
  mb?: number;      // 向下合并单元格数（不包括自身）
}
```

**globalConfig 参数：**
- `a`: 对齐方式
- `border`: 是否显示边框
- `w`: 默认列宽
- `h`: 默认行高
- `fg`: 背景色
- `c`: 文本颜色

**示例：生成会议记录模板**

```javascript
const t = [
  [
    { v: "Meeting Minutes", s: 16, mr: 3, fg: "#eee", h: 45, b: true },
    { w: 160 }, "", { w: 160 }
  ],
  ["Time", "", "Location", ""],
  ["Host", "", "Recorder", ""],
  ["Expected", "", "Present", ""],
  ["Absent Members", { mr: 2 }, "", ""],
  ["Topic", { mr: 2 }, "", ""],
  [{ v: "Content", h: 280 }, { mr: 2 }, "", ""],
  [{ v: "Remarks", h: 80 }, { mr: 2 }, "", ""]
]; // 必须是矩形矩阵

SN.activeSheet.insertTable(t, "A1", {
  border: true,
  a: "c",
  h: 35,
  w: 140
});
```

**注意**：
- 对于合并单元格（`mr`/`mb`），需要添加相同数量的空字符串占位符，保持二维数组的矩形结构。
- 例如：`{ mr: 2 }, "", ""`

### 图形对象

#### `addDrawing(config: object): Drawing`
添加图形对象（图表、图片等）。

**示例：添加图表**

```javascript
SN.activeSheet.addDrawing({
  type: 'chart',
  startCell: 'B2',
  option: {
    title: { text: '销售趋势图' },
    legend: {
      data: ['销量'] // 或使用引用: `${sheet.name}!B3`
    },
    xAxis: {
      type: 'category',
      data: ['一月', '二月', '三月'] // 或引用: `${sheet.name}!C2:E2`
    },
    yAxis: { type: 'value' },
    series: [
      {
        name: '销量',
        type: 'line',
        data: [820, 932, 901] // 或引用: `${sheet.name}!C3:E3`
      }
    ]
  }
});
```

#### `getDrawingsByCell(cellRef: CellRef): Drawing[]`
获取指定单元格位置的所有图形对象。

```javascript
const drawings = sheet.getDrawingsByCell("B2");
```

#### `removeDrawing(id: string): void`
删除指定 ID 的图形对象。

```javascript
sheet.removeDrawing("drawing-id");
```

---

## Cell

Cell 类代表单个单元格。

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `isMerged` | `boolean` | 是否为合并单元格 |
| `master` | `CellNum \| null` | 如果是合并单元格，指向主单元格 |
| `hyperlink` | `object` | 超链接：`{target, location, tooltip}` |
| `dataValidation` | `object` | 数据验证规则 |
| `editVal` | `string` | 编辑值或公式（可写） |
| `calcVal` | `any` | 计算值（只读） |
| `showVal` | `string` | 显示值（只读） |
| `numFmt` | `string` | 数字格式 |
| `font` | `object` | 字体样式 |
| `alignment` | `object` | 对齐方式 |
| `border` | `object` | 边框样式 |
| `fill` | `object` | 填充样式 |

### dataValidation 对象结构

```typescript
{
  type: string;              // 验证类型
  operator: string;          // 操作符
  allowBlank: boolean;       // 是否允许空白
  formula1: any;             // 公式1（type=list 时为数组）
  formula2: any;             // 公式2
  showInputMessage: boolean; // 是否显示输入提示
  promptTitle: string;       // 提示标题
  prompt: string;            // 提示内容
  showErrorMessage: boolean; // 是否显示错误提示
  errorTitle: string;        // 错误标题
  error: string;             // 错误内容
  errorStyle: string;        // 错误样式
  showDropDown: boolean;     // 是否显示下拉框
}
```

### 示例

```javascript
const cell = sheet.getCell(0, 0);

// 设置单元格值
cell.editVal = "Hello World";

// 设置公式
cell.editVal = "=SUM(A1:A10)";

// 读取显示值
console.log(cell.showVal);

// 设置样式
cell.fill = { fgColor: '#FFFF00' };
cell.font = { bold: true, size: 14 };
```

---

## Row

Row 类代表行，提供行级别的属性和操作。

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `cells` | `Cell[]` | 该行的所有单元格 |
| `height` | `number` | 行高（像素） |
| `hidden` | `boolean` | 是否隐藏 |
| `rIndex` | `number` | 行索引 |
| `numFmt` | `string` | 数字格式 |
| `font` | `object` | 字体样式 |
| `alignment` | `object` | 对齐方式 |
| `border` | `object` | 边框样式 |
| `fill` | `object` | 填充样式 |

### 方法

#### `getCell(c: number): Cell`
获取该行的指定列单元格。

```javascript
const row = sheet.getRow(0); // 获取第一行
const cell = row.getCell(0); // 获取该行第一列的单元格
```

### 示例

```javascript
const row = sheet.getRow(5);

// 设置行高
row.height = 30;

// 隐藏行
row.hidden = true;

// 设置行样式
row.fill = { fgColor: '#F0F0F0' };
row.font = { bold: true };
```

---

## Col

Col 类代表列，提供列级别的属性和操作。

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `cells` | `Cell[]` | 该列的所有单元格 |
| `hidden` | `boolean` | 是否隐藏 |
| `width` | `number` | 列宽（像素） |
| `cIndex` | `number` | 列索引 |
| `numFmt` | `string` | 数字格式 |
| `font` | `object` | 字体样式 |
| `alignment` | `object` | 对齐方式 |
| `border` | `object` | 边框样式 |
| `fill` | `object` | 填充样式 |

### 方法

#### `getCell(r: number): Cell`
获取该列的指定行单元格。

```javascript
const col = sheet.getCol(0); // 获取第一列（A列）
const cell = col.getCell(0); // 获取该列第一行的单元格
```

### 示例

```javascript
const col = sheet.getCol(2); // 获取 C 列

// 设置列宽
col.width = 120;

// 隐藏列
col.hidden = true;

// 设置列样式
col.fill = { fgColor: '#E0E0E0' };
col.alignment = { horizontal: 'center' };
```

---

## Drawing

Drawing 类代表图表、图片、连接线或形状等图形对象。

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `id` | `string` | 唯一标识符 |
| `type` | `string` | 类型：`chart`、`image`、`connector`、`shape` |
| `startCell` | `CellNum` | 起始单元格位置 |
| `offsetX` | `number` | X 轴偏移（默认 0） |
| `offsetY` | `number` | Y 轴偏移（默认 0） |
| `width` | `number` | 宽度（默认 480） |
| `height` | `number` | 高度（默认 288） |
| `option` | `object` | 图表配置（type=chart 时，与 ECharts 配置相同） |
| `imageBase64` | `string` | Base64 图片数据（type=image 时） |
| `updRender` | `boolean` | 是否更新渲染 |

### 方法

#### `updIndex(direction: string): void`
更新图层顺序。

参数值：`"up"` | `"down"` | `"top"` | `"bottom"`

```javascript
drawing.updIndex("top"); // 移到最上层
```

---

## Utils

Utils 类提供坐标转换等实用方法。

### 方法

#### `numToChar(num: number): string`
数字转字母列标。

```javascript
SN.Utils.numToChar(0); // "A"
SN.Utils.numToChar(25); // "Z"
SN.Utils.numToChar(26); // "AA"
```

#### `charToNum(char: string): number`
字母列标转数字。

```javascript
SN.Utils.charToNum("A"); // 0
SN.Utils.charToNum("Z"); // 25
SN.Utils.charToNum("AA"); // 26
```

#### `rangeNumToStr(rangeNum: RangeNum): string`
范围对象转字符串。

```javascript
SN.Utils.rangeNumToStr({s:{r:0,c:0}, e:{r:2,c:2}}); // "A1:C3"
```

#### `cellStrToNum(cellStr: string): CellNum`
单元格字符串转数字对象。

```javascript
SN.Utils.cellStrToNum("A1"); // {r:0, c:0}
```

#### `cellNumToStr(cellNum: CellNum): string`
单元格数字对象转字符串。

```javascript
SN.Utils.cellNumToStr({r:0, c:0}); // "A1"
```

---

## UndoRedoClass

管理操作历史，支持撤销和重做功能。

### 方法

#### `undo(): void`
撤销上一步操作。

```javascript
SN.UndoRedo.undo();
```

#### `redo(): void`
重做上一步操作。

```javascript
SN.UndoRedo.redo();
```

### 示例

```javascript
// 执行一些操作
sheet.getCell(0, 0).editVal = "Hello";
sheet.mergeCells("A1:B1");

// 撤销合并操作
SN.UndoRedo.undo();

// 撤销编辑操作
SN.UndoRedo.undo();

// 重做编辑操作
SN.UndoRedo.redo();
```

**注意**：
- 撤销/重做会自动记录大部分用户操作
- 通过 API 修改的内容也会被记录
- 历史记录栈有大小限制，过旧的操作会被清除
