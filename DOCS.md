# SheetNext 简介

SheetNext 是一个纯前端高性能 Excel 编辑器，只需数行代码即可集成到您的项目中。

## ✨ 核心特点

- **📊 完整的电子表格功能** - 支持单元格编辑、样式设置、公式引擎、图表绘制、数据排序、筛选等核心功能
- **🤖 AI 智能工作流** - 内置 AI 全自动操作流程，轻松实现模板生成、数据分析、公式编写、跨表逻辑操作等
- **📁 原生文件支持** - 原生支持 Excel (.xlsx)、CSV、JSON 文件的导入导出，无需额外插件
- **🚀 开箱即用** - 零配置开始使用，所有功能内置，无需单独安装依赖库
- **⚡ 高性能渲染** - 基于 Canvas 的虚拟滚动技术，轻松处理大数据量表格
- **🔄 快速迭代** - 版本持续更新，积极响应用户反馈和问题

## 📦 安装方式

SheetNext 提供多种安装方式，满足不同项目需求：

- **npm/yarn 安装** - 适用于现代前端项目（React、Vue、Angular 等）
- **浏览器直接引入** - 通过 CDN 或本地文件直接在 HTML 中使用

## 🔗 相关链接

- 🏠 [官网](https://www.sheetnext.com)
- 🎯 [在线体验](https://www.sheetnext.com/editor)
- 📦 [npm 包](https://www.npmjs.com/package/sheetnext)

# 快速开始

只需几行代码，即可将 SheetNext 集成到您的项目中。支持 npm 安装和浏览器直接引入两种方式，满足不同场景需求。

## 📦 使用 npm 安装

```bash
npm install sheetnext
```
```html
<div id="SNContainer" style="width:100vw;height:100vh;padding:0 7px 7px"></div>
```
```javascript
import SheetNext from 'sheetnext';
import 'sheetnext/dist/sheetnext.css';

// 注意设置容器#SNContainer宽高
const SN = new SheetNext(document.querySelector('#SNContainer'));
```

## 🌐 浏览器直接引入

```html
<!-- 引入样式 -->
<link rel="stylesheet" href="dist/sheetnext.css">

<!-- 编辑器容器 -->
<div id="SNContainer" style="width: 100vw; height: 100vh;padding:0 7px 7px"></div>

<!-- 引入脚本 -->
<!-- <script src="dist/sheetnext.umd.js"></script> -->

<!-- 初始化,注意设置宽高 -->
<!-- <script>
  const SN = new SheetNext(document.querySelector('#SNContainer'));
</script> -->
```

## ⚙️ 初始化配置

SheetNext 支持多种可选配置参数，用于定制编辑器的功能和外观。

```javascript
const SN = new SheetNext(document.querySelector('#container'), {
  AI_URL: "http://localhost:3000/sheetnextAI",  // AI 中转地址
  AI_TOKEN: "your-token",                        // AI 中转 token
  licenseKey: "your-license-key",                // 授权密钥
  menuList: (defaultList) => { /* ... */ },      // 自定义菜单栏
  menuRight: '<div>&copy SheetNext</div>'                // 菜单栏右侧自定义内容
});
```

## 配置参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `AI_URL` | `string` | AI 服务中转地址，用于配置 AI 功能的后端接口 |
| `AI_TOKEN` | `string` | AI 服务访问令牌，用于鉴权认证 |
| `licenseKey` | `string` | 商业版授权密钥，社区版可不填 |
| `menuList` | `function` | 自定义顶部菜单栏，接收默认菜单并返回修改后的菜单 |
| `menuRight` | `string` | 菜单栏右侧区域的自定义 HTML 内容 |

## menuList 自定义菜单示例

```javascript
const SN = new SheetNext(document.querySelector('#container'), {
  menuList: (defaultList) => {
    // 在"文件"菜单末尾添加自定义项
    defaultList[0].items.push({
      label: '我的自定义功能',
      handler: () => alert('这是自定义菜单项！')
    });

    // 添加新的顶级菜单
    defaultList.push({
      label: '帮助',
      items: [
        { label: '使用文档', handler: () => window.open('https://www.sheetnext.com/docs') },
        { label: '关于', handler: () => alert('SheetNext v1.0') }
      ]
    });

    return defaultList;
  }
});
```

**MenuList 结构定义：**

```typescript
interface MenuItem {
  label: string;              // 菜单项标签
  handler?: () => void;       // 点击处理函数
  disabled?: boolean;         // 是否禁用
  tip?: string;               // 提示信息（右侧显示）
  divider?: boolean;          // 是否为分隔线
}

interface Menu {
  label: string;              // 菜单标签
  items: MenuItem[];          // 菜单项列表
}

type MenuList = Menu[];
```

**注意事项：**
- `menuList` 和 `menuRight` 只能在初始化时配置，后续无法修改
- 如果不传入 `menuList`，将使用默认菜单（包含：文件、插入、公式、数据、视图、更多）
- AI 功能需要配置 `AI_URL` 才能使用，详见 AI 中转配置

# 工作簿级别

SheetNext 是主入口类，管理整个电子表格应用。

## 核心属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `workbookName` | `string` | 工作簿名称（可读写，长度1-255字符） |
| `activeSheet` | `Sheet` | 当前激活的工作表（可读写） |
| `sheets` | `Sheet[]` | 所有工作表数组 |
| `sheetNames` | `string[]` | 工作表名称列表（只读） |
| `containerDom` | `HTMLElement` | 编辑器容器元素 |
| `namespace` | `string` | 实例的全局命名空间（如 `SN_0`） |
| `locked` | `boolean` | 是否锁定工作表切换/操作 |

## 核心方法

### 新建工作表
`addSheet(name?: string): Sheet`

添加新工作表，名称可选（自动生成 Sheet1、Sheet2 等）。

```javascript
const newSheet = SN.addSheet("销售数据");
const autoSheet = SN.addSheet(); // 自动命名
```

**规则：** 名称不重复、长度1-31字符、不含特殊符号 `: / \ * ? [ ]`

### 删除工作表
`delSheet(name: string): void`

删除指定工作表（至少保留一个可见工作表）。

```javascript
SN.delSheet("Sheet2");
```

### 根据名称获取工作表
`getSheetByName(name: string): Sheet | null`

根据名称获取工作表。

```javascript
const sheet = SN.getSheetByName("Sheet1");
```

### 根据索引获取可见工作表
`getVisibleSheetByIndex(index: number): Sheet`

获取可见工作表（按索引，隐藏工作表不计入）。

```javascript
const firstSheet = SN.getVisibleSheetByIndex(0);
```

### 手动重新渲染
`r(): void`

手动触发画布重新渲染（批量修改后使用）。

```javascript
// 批量修改后刷新
for (let i = 0; i < 100; i++) {
  sheet.getCell(i, 0).editVal = i;
}
SN.r();
```

### 获取工作簿数据
`getData(): object`

获取完整的工作簿数据（JSON格式），包含所有工作表、单元格数据、样式、公式、图表等。

```javascript
// 获取工作簿数据
const data = SN.getData();
```

**使用场景：**
- 数据备份和恢复
- 数据持久化到数据库或本地存储
- 数据分析和处理
- 跨系统数据传输

**示例：保存到 localStorage**

```javascript
// 保存数据
const data = SN.getData();
localStorage.setItem('sheetData', JSON.stringify(data));

// 读取数据
const savedData = JSON.parse(localStorage.getItem('sheetData'));
SN.setData(savedData);
```

### 加载工作簿数据
`setData(data: object): boolean`

加载完整的工作簿数据，替换当前所有工作表内容。

```javascript
SN.setData(data):boolean;
```

### 导入文件
`import(file: File): Promise<void>`

导入文件，支持 `.xlsx`、`.csv` 和 `.json` 格式。

```javascript
// 通过文件选择器导入
const fileInput = document.createElement('input');
fileInput.type = 'file';
fileInput.accept = '.xlsx,.csv,.json';
fileInput.onchange = (e) => {
  const file = e.target.files[0];
  SN.import(file);
};
fileInput.click();
```

**支持格式：**
- `.xlsx` - Excel 工作簿（完整支持）
- `.csv` - 逗号分隔值文件（导入为单个工作表）
- `.json` - SheetNext JSON 格式（包含样式、公式、图表等）

### 从URL导入文件
`importFromUrl(url: String): Promise<void>`

通过在线地址导入 Excel 文件（.xlsx 格式）。

```javascript
await SN.importFromUrl('https://example.com/data.xlsx');
```

### 导出文件
`export(type: string): void`

导出电子表格，支持 `"XLSX"`、`"CSV"` 和 `"JSON"` 格式。

```javascript
SN.export('XLSX'); // 导出为 Excel 文件
```

**格式说明：**
- `XLSX` - Excel 工作簿格式，支持多工作表、样式、公式、图表等
- `CSV` - 纯文本格式，仅导出当前激活工作表的数据（不含样式）
- `JSON` - SheetNext 专用格式，完整保存工作簿结构、样式、公式、图表等，适合数据备份和快速加载

## 多实例支持

SheetNext 支持同一页面创建多个独立实例：

```javascript
const editor1 = new SheetNext(document.querySelector('#container1'));
const editor2 = new SheetNext(document.querySelector('#container2'));

console.log(editor1.namespace); // "SN_0"
console.log(editor2.namespace); // "SN_1"
```

# 工作表级别

Sheet 类代表一个工作表。

## 属性

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
| `rowCount` | `number` | 行数（只读） |
| `colCount` | `number` | 列数（只读） |
| `xSplit` | `number` | 冻结列数 |
| `ySplit` | `number` | 冻结行数 |
| `drawings` | `Drawing[]` | 图形对象列表（图表、图片等） |

## 基础方法

### 获取指定行
`getRow(r: number): Row`

获取指定行对象。

```javascript
const row = sheet.getRow(0); // 获取第一行
```

### 获取指定列
`getCol(c: number): Col`

获取指定列对象。

```javascript
const col = sheet.getCol(0); // 获取第一列（A列）
```

### 获取指定单元格
`getCell(r: number, c: number): Cell`

获取指定单元格。

```javascript
const cell = sheet.getCell(0, 0); // 获取 A1 单元格
```

### 通过字符串获取单元格
`getCellByStr(cellStr: string): Cell`

通过字符串引用获取单元格。

```javascript
const cell = sheet.getCellByStr("A1");
```

### 范围字符串转数字
`rangeStrToNum(rangeStr: string): RangeNum`

将字符串范围转换为数字范围对象。

```javascript
const rangeNum = sheet.rangeStrToNum("A1:C3");
// 返回: {s:{r:0,c:0}, e:{r:2,c:2}}
```

## 区域遍历

### 遍历区域
`eachArea(rangeRef: RangeRef, callback: (r, c, index) => void, reverse?: boolean): void`

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

## 行列操作

### 显示所有隐藏行
`showAllHidRows(): void`

显示所有隐藏的行。

```javascript
sheet.showAllHidRows();
```

### 显示所有隐藏列
`showAllHidCols(): void`

显示所有隐藏的列。

```javascript
sheet.showAllHidCols();
```

### 插入行
`addRows(startR: number, num: number): void`

在指定位置插入行。

```javascript
sheet.addRows(5, 3); // 在第 5 行位置插入 3 行
```

**注意**：多个同时调用时，应反向遍历以避免索引混乱。

### 插入列
`addCols(startC: number, num: number): void`

在指定位置插入列。

```javascript
sheet.addCols(2, 2); // 在第 2 列位置插入 2 列
```

### 删除行
`delRows(startR: number, num: number): void`

删除指定行。

```javascript
sheet.delRows(5, 3); // 删除从第 5 行开始的 3 行
```

### 删除列
`delCols(startC: number, num: number): void`

删除指定列。

```javascript
sheet.delCols(2, 2); // 删除从第 2 列开始的 2 列
```

## 合并单元格

### 合并单元格
`mergeCells(rangeRef: RangeRef): void`

合并指定区域的单元格。

```javascript
sheet.mergeCells("A1:C3");
// 或
sheet.mergeCells({s:{r:0,c:0}, e:{r:2,c:2}});
```

### 取消合并单元格
`unMergeCells(cellRef: CellRef): void`

取消合并（传入区域内任意单元格引用）。

```javascript
sheet.unMergeCells("A1"); // 取消包含 A1 的合并区域
```

## 排序

### 区域排序
`rangeSort(sortItems: SortItem[], range?: RangeRef): void`

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

## 批量插入数据

### 批量插入数据
`insertTable(data: (ICellConfig | string | number)[][], startCell: CellRef, globalConfig?: object): RangeNum`

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

## 图形对象

### 添加图形对象
`addDrawing(config: object): Drawing`

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

### 获取单元格的图形对象
`getDrawingsByCell(cellRef: CellRef): Drawing[]`

获取指定单元格位置的所有图形对象。

```javascript
const drawings = sheet.getDrawingsByCell("B2");
```

### 删除图形对象
`removeDrawing(id: string): void`

删除指定 ID 的图形对象。

```javascript
sheet.removeDrawing("drawing-id");
```

---

# 单元格级别

Cell 类代表单个单元格，提供完整的数据、样式、验证等功能。

## 核心属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `editVal` | `string` | 编辑值或公式（可读写） |
| `calcVal` | `any` | 计算值（只读） |
| `showVal` | `string` | 显示值（只读） |
| `type` | `string` | 单元格类型（只读）：`string/number/date/time/dateTime/boolean/error` |
| `isFormula` | `boolean` | 是否为公式（只读） |
| `isMerged` | `boolean` | 是否为合并单元格 |
| `master` | `CellNum \| null` | 如果是合并单元格，指向主单元格 |

## 样式属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `font` | `object` | 字体样式 |
| `alignment` | `object` | 对齐方式 |
| `border` | `object` | 边框样式 |
| `fill` | `object` | 填充样式 |
| `numFmt` | `string` | 数字格式 |

## 功能属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `hyperlink` | `object` | 超链接配置 |
| `dataValidation` | `object` | 数据验证规则 |
| `validData` | `boolean` | 数据验证结果（只读） |

## font 对象结构

字体样式配置对象。

```typescript
{
  name?: string;        // 字体名称，如 'Arial', '微软雅黑'
  size?: number;        // 字体大小，默认 11
  bold?: boolean;       // 是否粗体
  italic?: boolean;     // 是否斜体
  underline?: string;   // 下划线：'single' | 'double' | 'none'
  strike?: boolean;     // 是否删除线
  color?: string;       // 字体颜色，格式：'#RRGGBB'
  vertAlign?: string;   // 上下标：'superscript' | 'subscript'
  outline?: boolean;    // 是否轮廓（Mac专用）
  charset?: number;     // 字符集
}
```

**示例：**

```javascript
const cell = sheet.getCell(0, 0);

// 基础字体设置
cell.font = {
  name: '微软雅黑',
  size: 14,
  bold: true
};

// 带颜色和下划线
cell.font = {
  size: 12,
  color: '#FF0000',
  underline: 'single'
};

// 删除线
cell.font = { strike: true };
```

## alignment 对象结构

对齐方式配置对象。

```typescript
{
  horizontal?: string;  // 水平对齐：'left' | 'center' | 'right' | 'fill' | 'justify' | 'distributed'
  vertical?: string;    // 垂直对齐：'top' | 'middle' | 'bottom' | 'justify' | 'distributed'
  wrapText?: boolean;   // 是否自动换行
  indent?: number;      // 缩进级别（仅 horizontal='left'/'right' 时有效）
}
```

**示例：**

```javascript
// 水平居中，垂直居中
cell.alignment = {
  horizontal: 'center',
  vertical: 'middle'
};

// 自动换行
cell.alignment = { wrapText: true };

// 左对齐并缩进2级
cell.alignment = {
  horizontal: 'left',
  indent: 2
};

// 两端对齐
cell.alignment = {
  horizontal: 'justify',
  vertical: 'top'
};
```

## border 对象结构

边框样式配置对象，可单独设置四个方向的边框。

```typescript
{
  top?: {
    style: string;    // 边框样式
    color?: string;   // 边框颜色，格式：'#RRGGBB'，默认 '#000000'
  };
  right?: { style: string; color?: string; };
  bottom?: { style: string; color?: string; };
  left?: { style: string; color?: string; };
  diagonal?: { style: string; color?: string; };
}
```

**边框样式 (style)：**
- `thin` - 细线（最常用）
- `medium` - 中等
- `thick` - 粗线
- `dotted` - 点线
- `dashed` - 虚线
- `double` - 双线
- `hair` - 极细线
- `dashDot` - 点划线
- `dashDotDot` - 双点划线
- `mediumDashed` - 中等虚线
- `mediumDashDot` - 中等点划线
- `mediumDashDotDot` - 中等双点划线
- `slantDashDot` - 斜点划线

**示例：**

```javascript
// 设置所有边框
cell.border = {
  top: { style: 'thin' },
  right: { style: 'thin' },
  bottom: { style: 'thin' },
  left: { style: 'thin' }
};

// 设置单边框
cell.border = {
  bottom: { style: 'medium', color: '#FF0000' }
};

// 添加对角线
cell.border = {
  diagonal: { style: 'thin', color: '#0000FF' }
};

// 清空边框
cell.border = {};
```

## fill 对象结构

填充样式配置对象，支持纯色填充和渐变填充。

### 纯色填充

```typescript
{
  type: 'pattern';       // 填充类型
  pattern: string;       // 图案类型
  fgColor?: string;      // 前景色，格式：'#RRGGBB'
  bgColor?: string;      // 背景色，格式：'#RRGGBB'
}
```

**图案类型 (pattern)：**
- `solid` - 纯色（最常用）
- `darkGray` - 深灰
- `mediumGray` - 中灰
- `lightGray` - 浅灰
- `gray125` - 12.5% 灰
- `gray0625` - 6.25% 灰
- `darkHorizontal` - 深色横线
- `darkVertical` - 深色竖线
- `darkDown` - 深色斜线（左上到右下）
- `darkUp` - 深色斜线（左下到右上）
- `darkGrid` - 深色网格
- `darkTrellis` - 深色斜网格
- `lightHorizontal` - 浅色横线
- `lightVertical` - 浅色竖线
- `lightDown` - 浅色斜线（左上到右下）
- `lightUp` - 浅色斜线（左下到右上）
- `lightGrid` - 浅色网格
- `lightTrellis` - 浅色斜网格

### 渐变填充

```typescript
{
  type: 'gradient';        // 填充类型
  gradientType?: string;   // 渐变类型：'linear' | 'path'
  degree?: number;         // 线性渐变角度（0-360）
  left?: number;           // 左侧偏移（0-1）
  right?: number;          // 右侧偏移（0-1）
  top?: number;            // 顶部偏移（0-1）
  bottom?: number;         // 底部偏移（0-1）
  stops: Array<{           // 渐变色停止点
    position: number;      // 位置（0-1）
    color: string;         // 颜色，格式：'#RRGGBB'
  }>;
}
```

**示例：**

```javascript
// 纯色背景（最常用）
cell.fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: '#FFFF00'
};

// 图案填充
cell.fill = {
  type: 'pattern',
  pattern: 'lightGrid',
  fgColor: '#FF0000',
  bgColor: '#FFFFFF'
};

// 线性渐变（从红到蓝）
cell.fill = {
  type: 'gradient',
  gradientType: 'linear',
  degree: 90,
  stops: [
    { position: 0, color: '#FF0000' },
    { position: 1, color: '#0000FF' }
  ]
};

// 路径渐变
cell.fill = {
  type: 'gradient',
  gradientType: 'path',
  left: 0.5,
  right: 0.5,
  top: 0.5,
  bottom: 0.5,
  stops: [
    { position: 0, color: '#FFFFFF' },
    { position: 1, color: '#000000' }
  ]
};

// 清空填充
cell.fill = {};
```

## numFmt 数字格式

数字格式字符串，用于控制数值、日期、时间的显示格式。

### 常用格式

| 格式字符串 | 说明 | 示例 |
|-----------|------|------|
| `0` | 整数 | 1234 |
| `0.00` | 保留2位小数 | 1234.50 |
| `#,##0` | 千分位分隔符 | 1,234 |
| `#,##0.00` | 千分位+2位小数 | 1,234.50 |
| `0%` | 百分比 | 50% |
| `0.00%` | 百分比+2位小数 | 50.25% |
| `0.00E+00` | 科学计数法 | 1.23E+03 |
| `# ?/?` | 分数 | 1 1/4 |
| `¥#,##0.00` | 货币格式 | ¥1,234.50 |
| `yyyy/m/d` | 日期（年/月/日） | 2024/1/15 |
| `m/d/yyyy` | 日期（月/日/年） | 1/15/2024 |
| `yyyy-mm-dd` | 日期（年-月-日） | 2024-01-15 |
| `h:mm` | 时间（时:分） | 14:30 |
| `h:mm:ss` | 时间（时:分:秒） | 14:30:25 |
| `yyyy/m/d h:mm:ss` | 日期时间 | 2024/1/15 14:30:25 |
| `[Red]0.00` | 负数显示红色 | -12.34 |
| `0.00;[Red]-0.00` | 正数黑色，负数红色 | 12.34 或 -12.34 |

### 格式代码说明

**数字部分：**
- `0` - 占位符，不足补0
- `#` - 占位符，不足不显示
- `,` - 千分位分隔符
- `.` - 小数点
- `%` - 百分比
- `?` - 空格占位（用于对齐）

**日期部分：**
- `yyyy` - 四位年份（2024）
- `yy` - 两位年份（24）
- `m` - 月份（1-12）
- `mm` - 月份补0（01-12）
- `mmm` - 月份简写（Jan-Dec）
- `mmmm` - 月份全称（January-December）
- `d` - 日期（1-31）
- `dd` - 日期补0（01-31）
- `ddd` - 星期简写（Sun-Sat）
- `dddd` - 星期全称（Sunday-Saturday）

**时间部分：**
- `h` - 小时（0-23）
- `hh` - 小时补0（00-23）
- `m` - 分钟（0-59）
- `mm` - 分钟补0（00-59）
- `s` - 秒（0-59）
- `ss` - 秒补0（00-59）
- `AM/PM` - 12小时制

**示例：**

```javascript
const cell = sheet.getCell(0, 0);

// 数字格式
cell.editVal = 1234.5;
cell.numFmt = '#,##0.00';  // 显示：1,234.50

// 百分比
cell.editVal = 0.258;
cell.numFmt = '0.00%';     // 显示：25.80%

// 货币
cell.editVal = 1234.5;
cell.numFmt = '¥#,##0.00'; // 显示：¥1,234.50

// 日期
cell.editVal = '2024/1/15';
cell.numFmt = 'yyyy-mm-dd'; // 显示：2024-01-15

// 时间
cell.editVal = '14:30:25';
cell.numFmt = 'h:mm:ss';    // 显示：14:30:25

// 自定义格式（正数、负数、零、文本）
cell.numFmt = '0.00;[Red]-0.00;"零";@';

// 清除格式（恢复常规）
cell.numFmt = null;
```

## hyperlink 对象结构

超链接配置对象。

```typescript
{
  target?: string;    // 外部链接 URL（如 'https://example.com'）
  location?: string;  // 内部链接位置（如 'Sheet2!A1'）
  tooltip?: string;   // 鼠标悬停提示文本
}
```

**示例：**

```javascript
// 外部链接
cell.editVal = '访问官网';
cell.hyperlink = {
  target: 'https://www.sheetnext.com',
  tooltip: '点击访问 SheetNext 官网'
};

// 内部链接（跳转到其他工作表）
cell.editVal = '查看数据';
cell.hyperlink = {
  location: 'Sheet2!A1',
  tooltip: '跳转到 Sheet2 的 A1 单元格'
};

// 移除超链接
cell.hyperlink = {};
```

## dataValidation 对象结构

数据验证规则配置对象，用于限制单元格可输入的内容。

```typescript
{
  type: string;              // 验证类型：'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom'
  operator?: string;         // 操作符：'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
  allowBlank?: boolean;      // 是否允许空白，默认 false
  formula1: any;             // 公式1（type='list' 时为数组）
  formula2?: any;            // 公式2（范围验证时使用）
  showInputMessage?: boolean; // 是否显示输入提示，默认 true
  promptTitle?: string;      // 输入提示标题
  prompt?: string;           // 输入提示内容
  showErrorMessage?: boolean; // 是否显示错误提示，默认 true
  errorTitle?: string;       // 错误提示标题
  error?: string;            // 错误提示内容
  errorStyle?: string;       // 错误样式：'stop' | 'warning' | 'information'
  showDropDown?: boolean;    // 是否显示下拉框（type='list' 时），默认 true
}
```

**示例：**

```javascript
// 下拉列表
cell.dataValidation = {
  type: 'list',
  formula1: ['优秀', '良好', '及格', '不及格'],
  showDropDown: true,
  promptTitle: '请选择',
  prompt: '请从列表中选择一个等级'
};

// 整数范围（1-100）
cell.dataValidation = {
  type: 'whole',
  operator: 'between',
  formula1: 1,
  formula2: 100,
  errorTitle: '输入错误',
  error: '请输入 1 到 100 之间的整数'
};

// 小数验证（大于0）
cell.dataValidation = {
  type: 'decimal',
  operator: 'greaterThan',
  formula1: 0,
  errorTitle: '输入错误',
  error: '请输入大于 0 的数字'
};

// 日期范围
cell.dataValidation = {
  type: 'date',
  operator: 'between',
  formula1: '2024/1/1',
  formula2: '2024/12/31',
  errorTitle: '日期错误',
  error: '请输入 2024 年的日期'
};

// 时间验证
cell.dataValidation = {
  type: 'time',
  operator: 'between',
  formula1: '9:00:00',
  formula2: '18:00:00',
  errorTitle: '时间错误',
  error: '请输入工作时间（9:00-18:00）'
};

// 文本长度限制
cell.dataValidation = {
  type: 'textLength',
  operator: 'lessThanOrEqual',
  formula1: 20,
  errorTitle: '文本过长',
  error: '最多输入 20 个字符'
};

// 移除验证
cell.dataValidation = {};
```

# 行级别

Row 类代表行，提供行级别的属性和操作。

## 属性

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

## 方法

### 获取该行的单元格
`getCell(c: number): Cell`

获取该行的指定列单元格。

```javascript
const row = sheet.getRow(0); // 获取第一行
const cell = row.getCell(0); // 获取该行第一列的单元格
```

## 示例

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

# 列级别

Col 类代表列，提供列级别的属性和操作。

## 属性

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

## 方法

### 获取该列的单元格
`getCell(r: number): Cell`

获取该列的指定行单元格。

```javascript
const col = sheet.getCol(0); // 获取第一列（A列）
const cell = col.getCell(0); // 获取该列第一行的单元格
```

## 示例

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

# 图形级别

Drawing 类代表图表、图片或形状（包括连接线）等图形对象。

## 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `id` | `string` | 唯一标识符（只读） |
| `type` | `string` | 类型：`chart`、`image`、`shape` |
| `shapeType` | `string` | 形状类型（type=shape 时）：`rect`、`line`、`straightConnector1` 等 |
| `isConnector` | `boolean` | 是否为连接线（只读） |
| `shapeStyle` | `object` | 形状样式（type=shape 时）：`{fill, stroke, strokeWidth, startArrow, endArrow}` |
| `shapeText` | `string` | 形状内的文本（type=shape 时） |
| `startCell` | `CellNum` | 起始单元格位置 |
| `offsetX` | `number` | X 轴偏移（默认 5） |
| `offsetY` | `number` | Y 轴偏移（默认 5） |
| `width` | `number` | 宽度（默认：chart=460, shape=100, image=自动） |
| `height` | `number` | 高度（默认：chart=260, shape=100, image=自动） |
| `option` | `object` | 图表配置（type=chart 时，与 ECharts 配置相同） |
| `imageBase64` | `string` | Base64 图片数据（type=image 时） |
| `area` | `object` | 图形对象的实际覆盖区域（只读）：`{s:{r,c,offsetX,offsetY}, e:{r,c,offsetX,offsetY}}` |
| `anchorType` | `string` | 锚点类型：`twoCell`（随单元格移动+缩放）、`oneCell`（仅随移动）、`absolute`（固定） |
| `updRender` | `boolean` | 是否更新渲染 |

## 方法

### 更新图层顺序
`updIndex(direction: string): void`

更新图层顺序。

参数值：`"up"` | `"down"` | `"top"` | `"bottom"`

```javascript
drawing.updIndex("top"); // 移到最上层
```

# 布局管理

Layout 类管理编辑器的界面布局，包括菜单栏、工具栏、公式栏、Sheet 标签栏和 AI 聊天面板等界面元素的显示与隐藏。

**说明：** Layout 类由 SheetNext 自动创建，通过 `SN.Layout` 访问。菜单栏相关配置请参见 [快速开始 - 初始化配置](#快速开始)。

## 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `showMenuBar` | `boolean` | 是否显示菜单栏（可读写） |
| `showToolbar` | `boolean` | 是否显示工具栏（可读写） |
| `showFormulaBar` | `boolean` | 是否显示公式栏（可读写） |
| `showSheetTabBar` | `boolean` | 是否显示 Sheet 标签栏（可读写） |
| `showAIChat` | `boolean` | 是否显示 AI 聊天面板（可读写） |
| `showAIChatWindow` | `boolean` | 是否显示 AI 聊天小窗口模式（可读写） |
| `isSmallWindow` | `boolean` | 当前是否为小窗口模式（宽度 < 900px）（只读） |
| `menuConfig` | `object` | 菜单配置对象（只读） |

# 工具方法

Utils 类提供坐标转换等实用方法。

## 方法

### 数字转字母列标
`numToChar(num: number): string`

数字转字母列标。

```javascript
SN.Utils.numToChar(0); // "A"
SN.Utils.numToChar(25); // "Z"
SN.Utils.numToChar(26); // "AA"
```

### 字母列标转数字
`charToNum(char: string): number`

字母列标转数字。

```javascript
SN.Utils.charToNum("A"); // 0
SN.Utils.charToNum("Z"); // 25
SN.Utils.charToNum("AA"); // 26
```

### 范围对象转字符串
`rangeNumToStr(rangeNum: RangeNum): string`

范围对象转字符串。

```javascript
SN.Utils.rangeNumToStr({s:{r:0,c:0}, e:{r:2,c:2}}); // "A1:C3"
```

### 单元格字符串转数字对象
`cellStrToNum(cellStr: string): CellNum`

单元格字符串转数字对象。

```javascript
SN.Utils.cellStrToNum("A1"); // {r:0, c:0}
```

### 单元格数字对象转字符串
`cellNumToStr(cellNum: CellNum): string`

单元格数字对象转字符串。

```javascript
SN.Utils.cellNumToStr({r:0, c:0}); // "A1"
```

### 显示消息提示
`msg(message: string): void`

显示临时消息提示（3秒后自动消失）。

```javascript
SN.Utils.msg("操作成功！");
```

### 显示弹窗
`modal(options: object): Promise`

显示模态弹窗，返回 Promise（确定时 resolve，取消时 reject）。

```javascript
// 基础用法
SN.Utils.modal({
  title: '提示',
  content: '确定要删除吗？',
  confirmText: '确定',
  cancelText: '取消'
}).then(() => {
  console.log('用户点击了确定');
}).catch(() => {
  console.log('用户取消了');
});
```

# 历史记录

管理操作历史，支持撤销和重做功能。

## 方法

### 撤销操作
`undo(): void`

撤销上一步操作。

```javascript
SN.UndoRedo.undo();
```

### 重做操作
`redo(): void`

重做上一步操作。

```javascript
SN.UndoRedo.redo();
```

## 示例

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
- 历史记录栈有大小限制，过旧的操作会被清除

# AI 功能

通过 `SN.AI` 访问 AI 辅助功能

## 方法

### 监听 AI 请求状态
`listenRequestStatus(callback: Function): Function`

监听 AI 请求的 HTTP 状态码（如 200、401、500 等），返回取消监听的函数。

```javascript
// 添加监听器
const unsubscribe = SN.AI.listenRequestStatus((httpStatus) => {
  if (httpStatus === 200) {
    console.log('AI 请求成功');
  } else if (httpStatus === 401) {
    console.log('未授权，请检查 AI_TOKEN');
  }
});

// 取消监听
unsubscribe();
```

# AI 中转配置

写一个接口将前端传入的message消息分发给你想对接的大模型，然后在前端配置好接口地址即可开始工作！

## 功能说明

AI 服务中转层是连接 SheetNext 前端与大模型 API 的桥梁，主要负责以下核心功能：

**核心功能：**

1. **消息格式转换** - 将 SheetNext 提供的通用消息结构转换为目标大模型（如 Claude、GPT 等）所需的标准格式
2. **流式数据处理** - 实现 AI 响应的流式接收与转发，提升用户交互体验
3. **安全隔离** - 在服务端隐藏真实的 API Key，避免密钥泄露风险
4. **使用统计** - 企业可在中转层统计 Token 消耗、请求次数等关键数据

## 核心架构

```
┌─────────────┐         ┌──────────────┐         ┌─────────────┐
│             │         │              │         │             │
│  SheetNext  │────────▶│  中转服务器   │──────▶│各种大模型API │
│   前端      │  HTTP   │  (您的服务器) │  HTTPS  │  (Claude等) │
│             │◀────────│              │◀──────│             │
└─────────────┘  SSE流  └──────────────┘  Stream └─────────────┘
                                  │
                                  ▼
                          ┌──────────────┐
                          │ 使用统计/日志 │
                          └──────────────┘
```

**工作流程：**

1. **前端请求** - SheetNext 发送包含 `messages` 数组的 POST 请求到中转服务器
2. **格式转换** - 中转服务器将通用格式转换为目标大模型的专用格式
3. **API 调用** - 使用服务端存储的 API Key 调用大模型 API
4. **流式响应** - 接收大模型的流式响应，转换后通过 SSE (Server-Sent Events) 返回前端

## 完整示例

通用中转完整实现示例：

**安装依赖：**

```bash
npm install @anthropic-ai/sdk openai
```

**完整代码：**

```javascript
/**
 * SheetNext AI & claude/openai 中转服务器示例 Node.js 版本
 * 2025.10.17 v1.0.0
 */

const http = require('http');
const Anthropic = require('@anthropic-ai/sdk');
const OpenAI = require('openai');

// ======= 配置 =======
const CONFIG = {
    model: 'claude-sonnet-4-5-20250929', // 设置模型名称，自动判断使用 claude 还是 openai
    claude: {
        apiKey: 'your-apiKey',
        baseURL: 'https://xx.xx.xx/'
    },
    openai: {
        apiKey: 'your-apiKey',
        baseURL: 'https://xx.xx.xx/v1'
    }
};

const anthropic = new Anthropic({ apiKey: CONFIG.claude.apiKey, baseURL: CONFIG.claude.baseURL });
const openai = new OpenAI({ apiKey: CONFIG.openai.apiKey, baseURL: CONFIG.openai.baseURL });

// ======= message默认是openai格式，claude请求时转为它适配格式 =======
const convertToClaudeMessages = (messages) => {
    const system = [];
    const claudeMessages = [];
    let isFirstSystem = true;

    // 转换内容部分的辅助函数
    const convertContent = (content) => {
        const parts = Array.isArray(content) ? content : [{ type: 'text', text: content }];
        return parts.map(part => {
            if (part.type === 'text') {
                return { type: 'text', text: part.text };
            }
            if (part.type === 'image_url') {
                const [, mediaType, base64Data] = part.image_url.url.match(/data:(.*?);base64,(.*)/) || [];
                if (base64Data) {
                    return { type: 'image', source: { type: 'base64', media_type: mediaType || 'image/jpeg', data: base64Data } };
                }
            }
            return null;
        }).filter(Boolean);
    };

    for (const msg of messages) {
        if (msg.role === 'system') {
            if (isFirstSystem) {
                // 第一个 system：提取文本作为 system 参数（约定无图片）
                const text = typeof msg.content === 'string' ? msg.content : msg.content[0]?.text || '';
                if (text) system.push({ type: 'text', text });
                isFirstSystem = false;
            } else {
                // 其他 system：转为 user
                claudeMessages.push({ role: 'user', content: convertContent(msg.content) });
            }
        } else {
            // user/assistant 消息
            claudeMessages.push({ role: msg.role, content: convertContent(msg.content) });
        }
    }

    return { system, messages: claudeMessages };
};

// ======= Claude SDK =======
async function callClaudeSDK(messages, model, onChunk) {
    const { system, messages: claudeMessages } = convertToClaudeMessages(messages);

    // 打印请求结构（省略 base64 数据）
    const printableRequest = {
        system: system.map(s => s.type === 'image'
            ? { type: 'image', source: { ...s.source, data: `[${s.source.data?.length || 0} chars]` } }
            : s
        ),
        messages: claudeMessages.map(msg => ({
            role: msg.role,
            content: typeof msg.content === 'string' ? msg.content :
                msg.content.map(c => c.type === 'image'
                    ? { type: 'image', source: { ...c.source, data: `[${c.source.data?.length || 0} chars]` } }
                    : c
                )
        }))
    };

    const stream = await anthropic.messages.create({
        model: model,
        max_tokens: 8192,
        system,
        messages: claudeMessages,
        stream: true,
        thinking: { type: "enabled", budget_tokens: 2000 }
    });

    for await (const event of stream) {
        if (event.type === 'content_block_delta') {
            const { delta } = event;
            if (delta?.type === 'thinking_delta' && delta.thinking) {
                onChunk({ type: 'think', delta: delta.thinking });
            } else if (delta?.type === 'text_delta') {
                onChunk({ type: 'text', delta: delta.text });
            }
        }
    }
}

// ======= OpenAI SDK =======
async function callOpenAISDK(messages, model, onChunk) {
    const stream = await openai.chat.completions.create({
        model: model,
        messages: messages, // 直接使用 OpenAI 格式的 messages
        stream: true
    });

    for await (const chunk of stream) {
        const delta = chunk.choices[0]?.delta;
        if (delta?.content) {
            onChunk({ type: 'text', delta: delta.content });
        }
    }
}

// ======= HTTP 处理 =======
async function handleChat(messages, res) {
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
    });

    let ended = false;
    const write = (data) => !ended && !res.writableEnded && res.write(data);
    const onChunk = (chunk) => write(`data: ${JSON.stringify(chunk)}\n\n`);

    try {
        // 根据模型名称自动判断使用哪个 provider
        const provider = CONFIG.model.toLowerCase().includes('claude') ? 'claude' : 'openai';
        if (provider === 'openai') {
            await callOpenAISDK(messages, CONFIG.model, onChunk);
        } else {
            await callClaudeSDK(messages, CONFIG.model, onChunk);
        }
        write(`data: [DONE]\n\n`);
    } catch (error) {
        write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
    } finally {
        ended = true;
        res.end();
    }
}

// ======= HTTP 服务器 =======
http.createServer(async (req, res) => {
    const corsHeaders = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
    };

    if (req.method === 'OPTIONS') {
        res.writeHead(200, corsHeaders);
        return res.end();
    }

    if (req.url === '/sheetnextAI' && req.method === 'POST') {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', async () => {
            try {
                const { messages } = JSON.parse(body);
                if (!Array.isArray(messages)) throw new Error('Invalid messages');
                await handleChat(messages, res);
            } catch (error) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: error.message }));
            }
        });
    } else {
        res.writeHead(404);
        res.end('Not Found');
    }
}).listen(3000, () => console.log('🚀 Server running on http://localhost:3000'));
```

**配置说明：**

**判断规则:**
- 如果模型名称包含 `claude`(不区分大小写) → 使用 Claude SDK
- 其他情况 → 使用 OpenAI SDK

## 消息格式

**请求格式：**

SheetNext 发送的请求体格式：

```json
{
  "messages": [
    {
      "role": "system",
      "content": "你是一个电子表格助手..."
    },
    {
      "role": "user",
      "content": "帮我分析销售数据"
    },
    {
      "role": "assistant",
      "content": "好的，我来帮您分析..."
    },
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text": "某区域图片"
        },
        {
          "type": "image_url",
          "image_url": {
            "url": "data:image/png;base64,iVBORw0KGgoAAAANS..."
          }
        }
      ]
    }
  ]
}
```

**响应格式：**

您的服务器应该返回 SSE 流：

```
data: {"type":"text","delta":"我"}

data: {"type":"text","delta":"来"}

data: {"type":"text","delta":"帮"}

data: [DONE]
```
