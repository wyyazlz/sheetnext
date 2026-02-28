# Core API Details

> Last updated: 2026-02-28

## 📦 安装与初始化

### 使用 npm 安装

```bash
npm install sheetnext
```

```html
<!-- 放置编辑器的容器 -->
<div id="SNContainer" style="width:100vw;height:100vh;padding:0 7px 7px"></div>
```

```javascript
import SheetNext from 'sheetnext';
import 'sheetnext.css';

const SN = new SheetNext(document.querySelector('#SNContainer'));
```

### 浏览器直接引入（UMD）

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SheetNext Demo</title>
  <link rel="stylesheet" href="dist/sheetnext.css">
</head>
<body>
  <div id="SNContainer" style="width:100vw;height:100vh;padding:0 7px 7px"></div>
  <script src="dist/sheetnext.umd.js"></script>
  <script>
    const SN = new SheetNext(document.querySelector('#SNContainer'));
  </script>
</body>
</html>
```

## 📐 类型约定（开发约束）

> 本节作为 API 细节文档统一前置约束，调用前建议先阅读。

### 坐标与区域（0-based）

**字段含义**：`r` = row（行），`c` = column（列），`s` = start（起始），`e` = end（结束）

| 类型 | 说明 | 示例 |
| --- | --- | --- |
| CellNum | 单元格数字坐标 | `{r: 0, c: 0}` |
| CellStr | 单元格字符串地址 | `"A1"`、`"B2"` |
| CellRef | `CellNum | CellStr` | `{r:0,c:0}` 或 `"A1"` |
| RangeNum | 区域数字坐标 | `{s:{r:0,c:0}, e:{r:2,c:2}}` |
| RangeStr | 区域字符串地址 | `"A1:C3"`、`"A:B"`、`"1:2"` |
| RangeRef | `RangeNum | RangeStr` | 支持数字或字符串格式 |

> `r = 0` 表示第 1 行，`c = 0` 表示第 A 列。

### 样式对象 Style

- `font`: `name | size | color | bold | italic | underline | strike`
- `fill`: `type | pattern | fgColor`
- `alignment`: `horizontal | vertical | wrapText | textRotation | indent`
- `border`: `left | right | top | bottom`（每个方向含 `style` 与 `color`）
- `numFmt`: 数字格式字符串（如 `#,##0.00`、`yyyy-mm-dd`、`0.00%`）
- `protection`: `locked | hidden`

### 参数约束

- 优先使用文档声明类型，避免传入未定义字段。
- 涉及区域参数时，优先保证起始/结束坐标合法且边界有序。
- 涉及样式写入时，建议最小变更（仅传需要改动的字段）。

## 枚举类型

> 通过 `SN.Enum.XXX` 访问，或按需导入 `import { XXX } from 'sheetnext/enum'`

### EventPriority
- 事件优先级（越小越先执行）

| Key | Value | Description |
| --- | --- | --- |
| SYSTEM | -2000 | 系统内部 |
| PROTECTION | -1000 | 保护检查 |
| PLUGIN | -500 | 插件 |
| NORMAL | 0 | 默认 |
| LATE | 500 | 后置处理 |
| AUDIT | 1000 | 审计日志 |

### PasteMode
- 粘贴模式

| Value | Description |
| --- | --- |
| all | 全部（默认） |
| value | 仅值 |
| formula | 仅公式 |
| format | 仅格式 |
| noBorder | 无边框 |
| colWidth | 列宽 |
| comment | 批注 |
| validation | 数据验证 |
| transpose | 转置 |
| valueNumberFormat | 值和数字格式 |
| allMergeConditional | 全部使用源主题 |
| formulaNumberFormat | 公式和数字格式 |

### PasteOperation
- 运算粘贴模式

| Value | Description |
| --- | --- |
| none | 无 |
| add | 加 |
| sub | 减 |
| mul | 乘 |
| div | 除 |

### AnchorType
- 图纸锚点类型

| Value | Description |
| --- | --- |
| twoCell | 随单元格移动和缩放 |
| oneCell | 仅随单元格移动 |
| absolute | 固定位置 |

### DrawingType
- 图纸类型

| Value | Description |
| --- | --- |
| chart | - |
| image | - |
| shape | - |
| slicer | - |

### ValidationType
- 数据验证类型

| Value | Description |
| --- | --- |
| list | 下拉列表 |
| whole | 整数 |
| decimal | 小数 |
| date | 日期 |
| time | 时间 |
| textLength | 文本长度 |
| custom | 自定义公式 |

### ValidationOperator
- 数据验证操作符

| Value | Description |
| --- | --- |
| between | 介于 |
| notBetween | 未介于 |
| equal | 等于 |
| notEqual | 不等于 |
| greaterThan | 大于 |
| greaterThanOrEqual | 大于等于 |
| lessThan | 小于 |
| lessThanOrEqual | 小于等于 |

### FilterOperator
- 筛选条件操作符

| Value | Description |
| --- | --- |
| equal | 等于 |
| notEqual | 不等于 |
| greaterThan | 大于 |
| greaterThanOrEqual | 大于等于 |
| lessThan | 小于 |
| lessThanOrEqual | 小于等于 |
| beginsWith | 开头是 |
| endsWith | 结尾是 |
| contains | 包含 |
| notContains | 不包含 |

### CFType
- 条件格式类型

| Value | Description |
| --- | --- |
| cellIs | 单元格值 |
| containsText | 包含文本 |
| timePeriod | 发生日期 |
| top10 | 前/后 N 项 |
| aboveAverage | 高于/低于平均值 |
| duplicateValues | 重复值 |
| uniqueValues | 唯一值 |
| dataBar | 数据条 |
| colorScale | 色阶 |
| iconSet | 图标集 |
| expression | 公式 |

### SortOrder
- 排序方向

| Value | Description |
| --- | --- |
| asc | 升序 |
| desc | 降序 |

### HorizontalAlign
- 对齐方式 - 水平

| Value | Description |
| --- | --- |
| left | - |
| center | - |
| right | - |
| justify | 两端对齐 |
| fill | 填充 |
| distributed | 分散对齐 |

### VerticalAlign
- 对齐方式 - 垂直

| Value | Description |
| --- | --- |
| top | - |
| center | - |
| bottom | - |
| justify | - |
| distributed | - |

### BorderStyle
- 边框样式

| Value | Description |
| --- | --- |
| none | - |
| thin | - |
| medium | - |
| thick | - |
| dashed | - |
| dotted | - |
| double | - |

## AI/AI.js

### AI
- AI 助手模块 - Tool Calling 架构

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | Yes | - | SheetNext 主实例 |
| option | Object | Yes | - | AI 配置 |
| option.AI_URL | string | Yes | - | 中转服务器地址 |
| option.AI_TOKEN | string | No | - | 认证令牌 |

#### Methods
##### `chatInput(con)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| con | - | Yes | - | - |

##### `clearChat()`

##### `async conversation(p, joinChat = true)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| p | - | Yes | - | - |
| joinChat | boolean | No | true | - |

##### `handleFileChange(event)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | - | Yes | - | - |

##### `listenRequestStatus(callback)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | - | Yes | - | - |

##### `async screenshot(addresses = [])`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| addresses | Array | No | [] | - |

##### `sendInfo(info)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| info | - | Yes | - | - |

## AutoFilter/AutoFilter.js

### AutoFilter

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 所属工作表 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | 所属工作表 |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| ref | string \| null | get/set | No | 获取/设置默认（sheet）筛选范围 |
| sortState | {colIndex: number, order: 'asc' \| 'desc'} \| null | get/set | No | 默认（sheet）作用域排序状态 |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| activeScopeId | string | No | 当前激活作用域 ID |
| columns | Map<number, Object> | No | 默认（sheet）作用域筛选条件 |
| enabled | boolean | No | 默认（sheet）作用域是否启用 |
| hasActiveFilters | boolean | No | 默认（sheet）作用域是否有生效筛选 |
| headerRow | number | No | 默认（sheet）作用域表头行 |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} \| null | No | 默认（sheet）作用域范围 |

#### Methods
##### `buildXml(scopeId = this._defaultScopeId): Object | null`
- 构建指定作用域 XML；默认导出 sheet 作用域

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | this._defaultScopeId | - |

**Returns**
- Type: `Object | null`

##### `clear(scopeId = this._defaultScopeId)`
- 完全清除筛选（包含范围）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | this._defaultScopeId | - |

##### `clearAllFilters(scopeId = null)`
- 清除筛选条件（保留范围）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | null | - |

##### `clearColumnFilter(colIndex, scopeId = null)`
- 清除某列筛选

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

##### `getColumnFilter(colIndex, scopeId = null): Object | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `Object | null`
- Get filter config for column in scope.

##### `getColumnValues(colIndex, scopeId = null): Array<{value: any, text: string, count: number, isEmpty: boolean}>`
- 获取某列的所有唯一值（用于筛选面板）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

**Returns**
- Type: `Array<{value: any, text: string, count: number, isEmpty: boolean}>`

##### `getEnabledScopes(): Array<{scopeId:string,ownerType:string,ownerId:string | null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>`

**Returns**
- Type: `Array<{scopeId:string,ownerType:string,ownerId:string | null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>`
- List enabled scopes sorted by range.

##### `getScopeHeaderRow(scopeId = null): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `number`
- Get header row index of scope.

##### `getScopeRange(scopeId = null): {s:{r:number,c:number},e:{r:number,c:number}} | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `{s:{r:number,c:number},e:{r:number,c:number}} | null`
- Get range of scope.

##### `getSortOrder(colIndex, scopeId = null)`
- 获取排序状态

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

##### `getTableScopeId(tableId): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableId | string | Yes | - | - |

**Returns**
- Type: `string`
- Build table scope id from table id.

##### `hasActiveFiltersInScope(scopeId = null): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `boolean`
- Check whether scope has active column filters.

##### `hasAnyActiveFilters(): boolean`

**Returns**
- Type: `boolean`
- Check whether any scope has active filters.

##### `hasFilter(colIndex, scopeId = null)`
- 判断列是否有筛选

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

##### `isColumnInRange(colIndex, scopeId = null)`
- 判断某列是否在筛选范围内

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

##### `isHeaderRow(rowIndex, scopeId = null)`
- 判断某行是否是表头行

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

##### `isRowFilteredHidden(rowIndex): boolean`
- 行是否因筛选被隐藏

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

##### `isScopeEnabled(scopeId = null): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `boolean`
- Check whether scope range is enabled.

##### `parse(xmlObj)`
- 从 worksheet AutoFilter 解析默认作用域

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | - |

##### `parseScope(scopeId, autoFilterXml, options = {})`
- 将指定作用域的筛选状态解析到内存（主要用于 table 导入）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | Yes | - | - |
| autoFilterXml | Object | Yes | - | - |
| options | {range?: Object, restoreHiddenRowsFromSheet?: boolean, ownerType?: string, ownerId?: string} | No | {} | - |

##### `registerTableScope(table, options = {}): string | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| table | Object | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `string | null`
- Register or update filter scope for table.

##### `setActiveScope(scopeId): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | Yes | - | - |

**Returns**
- Type: `string`
- Set active scope id when scope exists.

##### `setColumnFilter(colIndex, filter, scopeId = null)`
- 设置列筛选条件

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| filter | Object | Yes | - | - |
| scopeId | string | No | null | - |

##### `setRange(range, scopeId = this._defaultScopeId)`
- 设置筛选范围（默认 sheet scope；可指定 scopeId）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} | Yes | - | - |
| scopeId | string | No | this._defaultScopeId | - |

##### `setSortState(colIndex, order, scopeId = null)`
- 设置排序状态

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| order | 'asc' \| 'desc' | Yes | - | - |
| scopeId | string | No | null | - |

##### `unregisterTableScope(tableId, options = {}): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableId | string | Yes | - | - |
| options | {silent?:boolean} | No | {} | - |

**Returns**
- Unregister table scope and refresh view state.

## Cell/Cell.js

### Cell
- 单元格类

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | 单元格 XML 对象 |
| row | Row | Yes | - | 行对象 |
| cIndex | number | Yes | - | 列索引 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| cIndex | number | No | cIndex | 列索引 |
| isMerged | boolean | No | false | 是否为合并单元格 |
| master | {r:number, c:number} \| null | No | null | 合并单元格主单元格引用 |
| row | Row | No | row | 所属行 |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| alignment | Object | get/set | No | 对齐方式 |
| border | Object | get/set | No | 边框样式 |
| dataValidation | Object \| null | get/set | No | 数据验证配置 |
| editVal | string | get/set | No | 编辑值或公式 |
| fill | Object | get/set | No | 填充样式 |
| font | Object | get/set | No | 字体样式 |
| hyperlink | Object \| null | get/set | No | 超链接配置 |
| numFmt | string \| undefined | get/set | No | 数字格式 |
| protection | {locked: boolean, hidden: boolean} | get/set | No | 单元格保护配置 |
| richText | Array \| null | get/set | No | 富文本 runs 数组 |
| style | Object | get/set | No | 单元格样式 |
| type | string | get/set | No | 单元格类型 |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| accountingData | Object \| undefined | No | 获取会计格式结构化数据（供渲染使用） |
| buildXml | Object | No | 构建单元格 XML |
| calcVal | any | No | 计算值 |
| horizontalAlign | string | No | 计算后的水平对齐方式 |
| isFormula | boolean | No | 是否为公式 |
| isLocked | boolean | No | 是否锁定 |
| isSpillRef | boolean | No | 是否是 spill 引用单元格 |
| isSpillSource | boolean | No | 是否是 spill 源单元格 |
| showVal | string | No | 显示值 |
| spillArray | Array[] \| null | No | 获取完整的 spill 数组结果 |
| validData | boolean | No | 数据验证结果 |
| verticalAlign | string | No | 计算后的垂直对齐方式 |

## CF/CF.js

### CF
- CF 条件格式管理器<br>负责解析、存储和管理 Excel 条件格式规则

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 所属工作表 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| rules | Array<CFRule> | No | [] | - |
| sheet | Sheet | No | sheet | - |

#### Methods
##### `add(config = {}): CFRule`
- 添加条件格式规则

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | No | {} | 规则配置 |
| config.rangeRef | string | Yes | - | 应用范围 "A1:D10" |
| config.type | string | Yes | - | 规则类型 (colorScale/dataBar/iconSet/cellIs/expression/top10/aboveAverage/duplicateValues/containsText/timePeriod/containsBlanks/containsErrors) |

**Returns**
- Type: `CFRule`
- 创建的规则

##### `clearAllCache(): void`
- 清除所有缓存

##### `clearRangeCache(sqref): void`
- 清除指定范围的缓存

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sqref | string | Yes | - | 范围 |

##### `get(index): CFRule | null`
- 获取规则

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| index | number | Yes | - | 规则索引 |

**Returns**
- Type: `CFRule | null`

##### `getAll(): Array<CFRule>`
- 获取所有规则

**Returns**
- Type: `Array<CFRule>`

##### `getDxfId(dxf): number`
- 获取dxf样式索引（用于导出）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dxf | Object | Yes | - | 差异化样式对象 |

**Returns**
- Type: `number`
- dxfId

##### `getDxfList(): Array`
- 获取dxf样式列表（用于导出styles.xml）

**Returns**
- Type: `Array`

##### `getFormat(r, c): Object | null`
- 获取单元格的条件格式

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | 行索引 |
| c | number | Yes | - | 列索引 |

**Returns**
- Type: `Object | null`
- 格式对象

##### `parse(xmlObj): void`
- 从 xmlObj 解析条件格式数据

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | Sheet 的 XML 对象 |

##### `remove(index): boolean`
- 删除条件格式规则

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| index | number | Yes | - | 规则索引 |

**Returns**
- Type: `boolean`

##### `toXmlObject(): Array`
- 导出为worksheet XML节点数组

**Returns**
- Type: `Array`

## Col/Col.js

### Col
- 列对象

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | 列 XML 数据 |
| sheet | Sheet | Yes | - | 所属工作表 |
| cIndex | number | Yes | - | 列索引 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| cIndex | number | No | cIndex | 列索引 |
| sheet | Sheet | No | sheet | 所属工作表 |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| alignment | Object | get/set | No | 对齐样式 |
| border | Object | get/set | No | 边框样式 |
| collapsed | boolean | get/set | No | 大纲折叠标记 |
| fill | Object | get/set | No | 填充样式 |
| font | Object | get/set | No | 字体样式 |
| hidden | boolean | get/set | No | 是否隐藏 |
| numFmt | string \| undefined | get/set | No | 数字格式 |
| outlineLevel | number | get/set | No | 大纲层级（0-7） |
| style | Object | get/set | No | 列样式 |
| width | number | get/set | No | 列宽（像素） |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| cells | Array<Cell> | No | 获取列内所有单元格 |

## Comment/Comment.js

### Comment
- Comment 批注管理器<br>高性能设计：Map索引（cellRef -> CommentItem）+ O(1)查找<br>架构参考：完全复刻 Drawing 的设计模式

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | - | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| map | Map<string, CommentItem> | No | new Map() | - |
| sheet | Sheet | No | sheet | - |

#### Methods
##### `add(config = {}): CommentItem`
- 添加批注

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | No | {} | {cellRef, author, text, visible, ...} |

**Returns**
- Type: `CommentItem`

##### `clearCache()`
- 清除所有批注的位置缓存（视图变化时调用）

##### `get(cellRef): CommentItem | null`
- 获取批注

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string | Yes | - | 单元格引用 "A1" |

**Returns**
- Type: `CommentItem | null`

##### `getAll(): CommentItem[]`
- 获取所有批注列表

**Returns**
- Type: `CommentItem[]`

##### `parse(xmlObj)`
- 从 xmlObj 解析批注数据（从 Excel 文件导入）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | Sheet 的 XML 对象 |

##### `remove(cellRef): boolean`
- 删除批注

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string | Yes | - | 单元格引用 "A1" |

**Returns**
- Type: `boolean`

##### `toXmlObject(): Object | null`
- 导出为XML对象（用于Excel导出）<br>将线程批注转换为普通批注格式

**Returns**
- Type: `Object | null`

## Drawing/Drawing.js

### Drawing
- Drawing manager - CRUD, XML parse & export

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| map | Map<string, DrawingItem> | No | new Map() | - |
| sheet | - | No | sheet | - |

#### Methods
##### `addChart(chartOption, options?): DrawingItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| chartOption | Object | Yes | - | - |
| options | DrawingOptions | No | - | - |

**Returns**
- Type: `DrawingItem`

##### `addImage(imageBase64, options?): DrawingItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| imageBase64 | string | Yes | - | - |
| options | DrawingOptions | No | - | - |

**Returns**
- Type: `DrawingItem`

##### `addShape(shapeType, options?): DrawingItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| shapeType | string | Yes | - | - |
| options | DrawingOptions | No | - | - |

**Returns**
- Type: `DrawingItem`

##### `get(id): DrawingItem | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

**Returns**
- Type: `DrawingItem | null`

##### `getAll(): DrawingItem[]`

**Returns**
- Type: `DrawingItem[]`

##### `getByCell(cellRef): DrawingItem[]`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | Object \| string | Yes | - | - |

**Returns**
- Type: `DrawingItem[]`

##### `remove(id)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

## Formula/Formula.js

### Formula
- 公式计算器

## IO/IO.js

### IO

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | import("../Workbook/Workbook.js").default | Yes | - | - |

#### Methods
##### `export(type)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| type | 'XLSX' \| 'CSV' \| 'JSON' \| 'HTML' | Yes | - | - |

##### `exportAllImage()`

##### `getData(): Promise<Object>`

**Returns**
- Type: `Promise<Object>`

##### `async import(file): Promise<void>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| file | File | Yes | - | - |

**Returns**
- Type: `Promise<void>`
- Import from xlsx/csv/json file.

##### `async importFromUrl(fileUrl): Promise<void>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fileUrl | string | Yes | - | - |

**Returns**
- Type: `Promise<void>`

##### `setData(data): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| data | Object | Yes | - | - |

**Returns**
- Type: `boolean`

## Layout/Layout.js

### Layout
- 布局与工具栏管理

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | Yes | - | SheetNext instance. |
| options | Object | No | {} | Layout options. |
| options.menuRight | function | No | - | Callback `(defaultHTML: string) => string` to customize the top-right menu HTML. |
| options.menuList | function | No | - | Callback `(config: Array) => Array` to customize menu tab config. |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| chatWindowOffsetX | number | No | 0 | 聊天窗口拖拽偏移 X |
| chatWindowOffsetY | number | No | 0 | 聊天窗口拖拽偏移 Y |
| menuConfig | Object | No | MenuConfig.initDefaultMenuConfig(SN, options) | 菜单配置 |
| myModal | Object \| null | No | null | 通用弹窗实例 |
| SN | Object | No | SN | SheetNext 主实例 |
| StateSync | StateSync | No | new StateSync(SN) | 状态同步器 |
| toast | Object \| null | No | null | Toast 实例 |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| showAIChat | boolean | get/set | No | 是否显示 AI 聊天入口 |
| showAIChatWindow | boolean | get/set | No | 是否显示 AI 聊天窗口 |
| showFormulaBar | boolean | get/set | No | 是否显示公式栏 |
| showMenuBar | boolean | get/set | No | 是否显示菜单栏 |
| showPivotPanel | boolean | get/set | No | 是否显示透视表字段面板 |
| showSheetTabBar | boolean | get/set | No | 是否显示工作表标签栏 |
| showStats | boolean | get/set | No | 是否显示统计栏 |
| showToolbar | boolean | get/set | No | 是否显示工具栏 |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| isSmallWindow | boolean | No | 是否为小窗口 |
| isToolbarModeLocked | - | No | - |
| minimalToolbarTextEnabled | - | No | - |
| toolbarMode | - | No | - |

#### Methods
##### `autoOpenPivotPanel(pt): void`
- 自动打开透视表字段面板（点击透视表区域时调用）<br>如果用户手动关闭过，则不自动打开

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pt | PivotTable | Yes | - | 透视表实例 |

##### `closePivotPanel(): void`
- 关闭透视表字段面板（点击透视表区域外时调用）<br>不会设置手动关闭标记

##### `handleAllResize()`
- 处理所有resize相关的逻辑

##### `initColorComponents(scope?)`
- Initialize color components.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scope | Element | No | - | Optional scope. |

##### `initDragEvents()`
- 初始化拖拽事件

##### `initEventListeners()`
- Initialize all event listeners.

##### `initResizeObserver()`
- 初始化ResizeObserver

##### `openPivotPanel(pt): void`
- 打开透视表字段面板（用户手动调用，如工具栏按钮）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pt | PivotTable | Yes | - | 透视表实例 |

##### `refreshActiveButtons(): void`
- 刷新工具栏中带有 active getter 的按钮状态<br>用于格式刷等切换按钮的双向绑定

##### `refreshToolbar(): void`
- 统一刷新工具栏状态（面板切换、API修改后调用）<br>整合：样式状态 + checkbox + active按钮

##### `removeLoading()`
- Remove loading animation.

##### `scrollSheetTabs(dir): void`
- 滚动工作表标签

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dir | number | Yes | - | 滚动方向（1 或 -1） |

##### `throttleResize()`
- 节流处理resize事件

##### `toggleMinimalToolbar(enabled)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| enabled | - | Yes | - | - |

##### `toggleMinimalToolbarText(enabled)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| enabled | - | Yes | - | - |

##### `togglePivotPanel(pt): boolean`
- 切换透视表字段面板显示

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pt | PivotTable | Yes | - | 透视表实例 |

**Returns**
- Type: `boolean`
- 是否显示

##### `updateCanvasSize()`
- 更新Canvas尺寸

##### `updateContextualToolbar(context): void`
- 更新上下文工具栏（根据选择内容显示/隐藏上下文选项卡）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| context | Object | Yes | - | 上下文信息 { type: 'table'\|'pivotTable'\|'chart'\|null, data: any, autoSwitch: boolean }<br>  - autoSwitch: 是否自动切换到上下文选项卡（默认false，只显示不切换） |

## License/License.js

### License
- License 授权管理类 (Ed25519)

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetNext | Object | Yes | - | SheetNext 主实例 |
| licenseKey | string | Yes | - | 授权码 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| publicKey | string | No | 'ISuSucIeA0p8Lftf0YHGogmXqDFvLtEFE46XsoB1R9s=' | Ed25519 公钥（32 字节 Base64） |
| sheetNext | Object | No | sheetNext | SheetNext 主实例 |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| isActivated | boolean | No | 是否已激活授权 |
| isServiceActive | boolean | No | 服务包是否在有效期内 |

#### Methods
##### `async activate(licenseKey): Promise<void>`
- 激活授权

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| licenseKey | string | Yes | - | 授权码 |

**Returns**
- Type: `Promise<void>`

##### `base64Decode(str): string`
- Base64 解码（支持 UTF-8）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| str | string | Yes | - | Base64 字符串 |

**Returns**
- Type: `string`

##### `getInfo(): {domains: string[], serviceExpireDate: string, daysLeft: number, isServiceActive: boolean} | null`

**Returns**
- Type: `{domains: string[], serviceExpireDate: string, daysLeft: number, isServiceActive: boolean} | null`

##### `async verifySignature(data, signatureBase64): Promise<boolean>`
- 验证 Ed25519 签名（使用 Web Crypto API）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| data | string | Yes | - | 原始数据字符串 |
| signatureBase64 | string | Yes | - | Base64 签名 |

**Returns**
- Type: `Promise<boolean>`

## PivotTable/PivotTable.js

### PivotTable
- 透视表管理器<br>负责透视表的增删改查

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| map | Map<string, PivotTableItem> | No | new Map() | - |
| sheet | Sheet | No | sheet | - |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| size | number | No | - |

#### Methods
##### `add(config): PivotTable`
- 添加透视表

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | 配置对象 |
| config.sourceSheet | Sheet | Yes | - | 源工作表 |
| config.sourceRangeRef | string | Yes | - | 源数据范围 |
| config.cellRef | string \| Object | Yes | - | 目标单元格 |

**Returns**
- Type: `PivotTable`

##### `forEach(callback)`
- 遍历所有透视表

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

##### `get(name): PivotTable | null`
- 获取透视表

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `PivotTable | null`

##### `getAll(): Array<PivotTable>`
- 获取所有透视表

**Returns**
- Type: `Array<PivotTable>`

##### `refreshAll()`
- 刷新所有透视表

##### `remove(name)`
- 删除透视表

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | 透视表名称 |

## Print/Print.js

### Print
- 打印设置与渲染

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | Yes | - | SheetNext 主实例 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | No | SN | SheetNext 主实例 |

#### Methods
##### `applySettingsToWorksheet(sheet, ws): void`
- 应用打印设置到 worksheet XML

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| ws | Object | Yes | - | worksheet XML 对象 |

##### `bindPreviewEvents(bodyEl, pages, layout): void`
- 绑定预览交互事件

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| bodyEl | HTMLElement | Yes | - | 弹窗 body |
| pages | Array<Object> | Yes | - | 页面数据 |
| layout | Object | Yes | - | 布局信息 |

##### `buildCellAddress(sheet, cell): string`
- 构建带工作表名的单元格地址

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| cell | {r: number, c: number} | Yes | - | 单元格坐标 |

**Returns**
- Type: `string`

##### `buildLayout(sheet, settings, options = {}): Object`
- 构建分页布局信息

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| settings | Object | Yes | - | 打印设置 |
| options | Object | No | {} | - |

**Returns**
- Type: `Object`

##### `buildPreviewHtml(pages, layout): string`
- 构建预览 HTML

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pages | Array<Object> | Yes | - | 页面数据 |
| layout | Object | Yes | - | 布局信息 |

**Returns**
- Type: `string`

##### `buildPreviewPageHtml(page, layout, pageIndex = 0, pageCount = 1): string`
- 构建单页预览 HTML

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| page | Object | Yes | - | 页面数据 |
| layout | Object | Yes | - | 布局信息 |
| pageIndex | number | No | 0 | - |
| pageCount | number | No | 1 | - |

**Returns**
- Type: `string`

##### `buildPrintHtml(pages, layout): string`
- 构建打印 HTML

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pages | Array<Object> | Yes | - | 页面数据 |
| layout | Object | Yes | - | 布局信息 |

**Returns**
- Type: `string`

##### `buildSegments(start, end, sizeGetter, maxSize, forcedStarts = []): Array<{s: number, e: number}>`
- 构建分页区段

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| start | number | Yes | - | 起始索引 |
| end | number | Yes | - | 结束索引 |
| sizeGetter | Function | Yes | - | 获取尺寸方法 |
| maxSize | number | Yes | - | 最大尺寸 |
| forcedStarts | Array | No | [] | - |

**Returns**
- Type: `Array<{s: number, e: number}>`

##### `async captureRange(sheet, range, width, height, settings): Promise<string>`
- 截图指定范围

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} | Yes | - | 范围 |
| width | number | Yes | - | 宽度 |
| height | number | Yes | - | 高度 |
| settings | Object | Yes | - | 打印设置 |

**Returns**
- Type: `Promise<string>`

##### `clearPrintArea(sheet = this.SN.activeSheet): void`
- 清空打印区域

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

##### `createDefaultSettings(): Object`
- 创建默认打印设置

**Returns**
- Type: `Object`

##### `ensureSettings(sheet = this.SN.activeSheet): Object | null`
- 获取并确保工作表的打印设置

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `getColSize(sheet, c): number`
- 获取列宽（像素）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| c | number | Yes | - | 列索引 |

**Returns**
- Type: `number`

##### `getMarginPixels(margins = DEFAULT_MARGINS): {left: number, right: number, top: number, bottom: number}`
- 计算边距像素值

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| margins | Object | No | DEFAULT_MARGINS | 边距配置 |

**Returns**
- Type: `{left: number, right: number, top: number, bottom: number}`

##### `getPageBreakGuides(sheet = this.SN.activeSheet): {range:Object,rowBreaks:Array<{index:number,manual:boolean}>,colBreaks:Array<{index:number,manual:boolean}>} | null`
- 获取分页辅助线信息（用于画布显示）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `{range:Object,rowBreaks:Array<{index:number,manual:boolean}>,colBreaks:Array<{index:number,manual:boolean}>} | null`

##### `getPaperDefinition(settings): Object`
- 获取纸张定义

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| settings | Object | Yes | - | 打印设置 |

**Returns**
- Type: `Object`

##### `getPrintRange(sheet, settings): {s: {r: number, c: number}, e: {r: number, c: number}}`
- 获取实际打印范围

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| settings | Object | Yes | - | 打印设置 |

**Returns**
- Type: `{s: {r: number, c: number}, e: {r: number, c: number}}`

##### `getRangeSize(sheet, range, includeHeadings): {width: number, height: number}`
- 获取范围尺寸

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} | Yes | - | 范围 |
| includeHeadings | boolean | Yes | - | 是否包含行列头 |

**Returns**
- Type: `{width: number, height: number}`

##### `getRowSize(sheet, r): number`
- 获取行高（像素）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| r | number | Yes | - | 行索引 |

**Returns**
- Type: `number`

##### `getSettings(sheet = this.SN.activeSheet): Object | null`
- 获取打印设置（深拷贝）

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `getUsedRange(sheet): {s: {r: number, c: number}, e: {r: number, c: number}}`
- 获取已使用区域

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |

**Returns**
- Type: `{s: {r: number, c: number}, e: {r: number, c: number}}`

##### `initSheetPrintSettings(sheet): Object`
- 初始化工作表打印设置

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |

**Returns**
- Type: `Object`

##### `insertPageBreak(cell = sheet.activeCell, sheet = this.SN.activeSheet): {row:boolean, col:boolean} | null`
- 在当前活动单元格插入手动分页符

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | {r:number, c:number} | No | sheet.activeCell | 单元格位置 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `{row:boolean, col:boolean} | null`

##### `normalizeRange(range, sheet): {s: {r: number, c: number}, e: {r: number, c: number}} | null`
- 规范化范围对象

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| null | Yes | - | 范围 |
| sheet | Sheet | Yes | - | 工作表 |

**Returns**
- Type: `{s: {r: number, c: number}, e: {r: number, c: number}} | null`

##### `async preview(options = {}): Promise<void>`
- 打印预览

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | 预览选项 |

**Returns**
- Type: `Promise<void>`

##### `async print(options = {}): Promise<void>`
- 打印

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | 打印选项 |

**Returns**
- Type: `Promise<void>`

##### `readSettingsFromXml(xmlObj, sheet = null): Object`
- 从 XML 读取打印设置

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | worksheet XML 对象 |
| sheet | null | No | null | - |

**Returns**
- Type: `Object`

##### `async renderPages(sheet, settings, layout): Promise<Array<Object>>`
- 渲染所有页面

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 工作表 |
| settings | Object | Yes | - | 打印设置 |
| layout | Object | Yes | - | 布局信息 |

**Returns**
- Type: `Promise<Array<Object>>`

##### `setPageMargins(margins, sheet = this.SN.activeSheet): Object | null`
- 设置页边距

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| margins | Object | Yes | - | 边距配置 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `setPageSetup(setup, sheet = this.SN.activeSheet): Object | null`
- 设置页面配置

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| setup | Object | Yes | - | 页面设置 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `setPrintArea(range, sheet = this.SN.activeSheet): Object | null`
- 设置打印区域

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | 区域 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `setPrintOptions(options, sheet = this.SN.activeSheet): Object | null`
- 设置打印选项

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | Yes | - | 打印选项 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `setPrintTitles(titles, sheet = this.SN.activeSheet): void`
- 设置打印标题行/列

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| titles | {rows?: string \| number[] \| {s:number,e:number} \| null, cols?: string \| number[] \| {s:number,e:number} \| null} | Yes | - | 标题配置 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

##### `setSettings(partial, sheet = this.SN.activeSheet): Object | null`
- 合并更新打印设置

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| partial | Object | Yes | - | 部分设置 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

**Returns**
- Type: `Object | null`

##### `toggleGridlines(on, sheet = this.SN.activeSheet): void`
- 切换打印网格线

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| on | boolean | Yes | - | 是否启用 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

##### `toggleHeadings(on, sheet = this.SN.activeSheet): void`
- 切换打印行列头

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| on | boolean | Yes | - | 是否启用 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

##### `toggleShowPageBreaks(on, sheet = this.SN.activeSheet): void`
- 显示/隐藏分页符辅助线

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| on | boolean | Yes | - | 是否显示 |
| sheet | Sheet | No | this.SN.activeSheet | 工作表 |

## Row/Row.js

### Row
- 行对象

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | 行 XML 数据 |
| sheet | Sheet | Yes | - | 所属工作表 |
| rIndex | number | Yes | - | 行索引 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| cells | Array<Cell> | No | [] | 行内单元格数组 |
| rIndex | number | No | rIndex | 行索引 |
| sheet | Sheet | No | sheet | 所属工作表 |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| alignment | Object | get/set | No | 对齐样式 |
| border | Object | get/set | No | 边框样式 |
| collapsed | boolean | get/set | No | 大纲折叠标记 |
| fill | Object | get/set | No | 填充样式 |
| font | Object | get/set | No | 字体样式 |
| height | number | get/set | No | 行高（像素） |
| hidden | boolean | get/set | No | 是否隐藏 |
| numFmt | string \| undefined | get/set | No | 数字格式 |
| outlineLevel | number | get/set | No | 大纲层级（0-7） |
| style | Object | get/set | No | 行样式 |

#### Methods
##### `init(): void`
- 初始化行内单元格

## Sheet/Sheet.js

### Sheet

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| meta | Object | Yes | - | - |
| SN | SheetNext | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| AutoFilter | AutoFilter | No | new AutoFilter(this) | - |
| Canvas | - | No | SN.Canvas | - |
| CF | null | No | null | - |
| cols | Col[] | No | [] | - |
| Comment | Comment | No | new Comment(this) | - |
| Drawing | null | No | null | - |
| initialized | boolean | No | false | - |
| merges | RangeNum[] | No | [] | - |
| outlinePr | Object | No | { | - |
| PivotTable | null | No | null | - |
| printSettings | Object | No | null | - |
| protection | Object | No | new SheetProtection(this) | - |
| rId | string | No | meta['_$r:id'] | - |
| rows | Row[] | No | [] | - |
| showGridLines | boolean | No | true | - |
| showPageBreaks | boolean | No | false | - |
| showRowColHeaders | boolean | No | true | - |
| Slicer | null | No | null | - |
| SN | - | No | SN | - |
| Sparkline | null | No | null | - |
| Table | Table | No | new Table(this) | - |
| Utils | - | No | SN.Utils | - |
| vi | Object | No | null | - |
| views | Object[] | No | [{ pane: {} }] | - |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| activeAreas | RangeNum[] | get/set | No | - |
| activeCell | CellNum | get/set | No | - |
| defaultColWidth | number | get/set | No | - |
| defaultRowHeight | number | get/set | No | - |
| frozenCols | number | get/set | No | - |
| frozenRows | number | get/set | No | - |
| headHeight | number | get/set | No | - |
| hidden | boolean | get/set | No | - |
| indexWidth | number | get/set | No | - |
| name | string | get/set | No | - |
| viewStart | CellNum | get/set | No | - |
| zoom | number | get/set | No | - |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| colCount | number | No | - |
| rowCount | number | No | - |

#### Methods
##### `addCols(c, number = 1)`
- 插入列

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | Number | Yes | - | 插入位置的列索引 |
| number | Number | No | 1 | 插入的列数 |

##### `addRows(r, number = 1)`
- 插入行

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | Number | Yes | - | 插入位置的行索引 |
| number | Number | No | 1 | 插入的行数 |

##### `applyBrush(targetArea): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetArea | RangeNum | Yes | - | - |

**Returns**
- Type: `boolean`

##### `areaHaveMerge(area): Boolean`
- 检测区域中是否存在合并的单元格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | 区域对象 |

**Returns**
- Type: `Boolean`
- 是否存在合并单元格

##### `areasBorder(position, options, area = this.activeAreas)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| position | string | Yes | - | - |
| options | Object \| null | Yes | - | - |
| area | Array | No | this.activeAreas | - |

##### `cancelBrush(): boolean`

**Returns**
- Type: `boolean`

##### `clearClipboard()`
- 清除剪贴板

##### `clearDataValidation(range)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |

##### `consolidate(ranges, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ranges | - | Yes | - | - |
| options | Object | No | {} | - |

##### `copy(area = null, isCut = false)`
- 复制区域数据

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | No | null | 复制区域 { s: {r, c}, e: {r, c} } |
| isCut | boolean | No | false | 是否剪切 |

##### `cut(area = null)`
- 剪切区域数据

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | No | null | 剪切区域 |

##### `delCols(c, number = 1)`
- 删除列

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | Number | Yes | - | 删除起始位置的列索引 |
| number | Number | No | 1 | 删除的列数 |

##### `delRows(r, number = 1)`
- 删除行

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | Number | Yes | - | 删除起始位置的行索引 |
| number | Number | No | 1 | 删除的行数 |

##### `eachCells(ranges, callback, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ranges | RangeRef \| RangeRef[] | Yes | - | - |
| callback | (r:number,c:number,area:RangeNum)=>void | Yes | - | - |
| options | {reverse?:boolean,sparse?:boolean} | No | {} | - |

##### `flashFill(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `getAreaInviewInfo(area): Object`
- 获取区域在可见视图中的信息

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | 区域对象 |

**Returns**
- Type: `Object`
- 包含位置和尺寸信息的对象

##### `getCell(r, c?): Cell`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number \| string | Yes | - | - |
| c | number | No | - | - |

**Returns**
- Type: `Cell`

##### `getCellInViewInfo(rowIndex, colIndex, posMerge = true): Object`
- 获取单元格在可见视图中的信息

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | Number | Yes | - | 行索引 |
| colIndex | Number | Yes | - | 列索引 |
| posMerge | Boolean | No | true | 是否处理合并单元格 |

**Returns**
- Type: `Object`
- 包含位置和尺寸信息的对象

##### `getClipboardData()`
- 获取当前剪贴板数据

##### `getCol(c): Col`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | number | Yes | - | - |

**Returns**
- Type: `Col`

##### `getColIndexByScrollLeft(scrollLeft): Number`
- 根据水平滚动像素位置找到对应列索引

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollLeft | Number | Yes | - | 滚动像素位置 |

**Returns**
- Type: `Number`
- 列索引

##### `getRow(r): Row`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | - |

**Returns**
- Type: `Row`

##### `getRowIndexByScrollTop(scrollTop): Number`
- 根据垂直滚动像素位置找到对应行索引

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollTop | Number | Yes | - | 滚动像素位置 |

**Returns**
- Type: `Number`
- 行索引

##### `getScrollLeft(colIndex): Number`
- 获取指定列之前所有列的宽度总和

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | Number | Yes | - | 列索引 |

**Returns**
- Type: `Number`
- 宽度总和

##### `getScrollTop(rowIndex): Number`
- 获取指定行之前所有行的高度总和

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | Number | Yes | - | 行索引 |

**Returns**
- Type: `Number`
- 高度总和

##### `getTotalHeight(): Number`
- 获取所有行的总高度（缓存优化）

**Returns**
- Type: `Number`
- 总高度

##### `getTotalWidth(): Number`
- 获取所有列的总宽度（缓存优化）

**Returns**
- Type: `Number`
- 总宽度

##### `groupCols(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `groupRows(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `hasClipboardData()`
- 检查是否有剪贴板数据

##### `hyperlinkJump()`
- Jump to hyperlink target of active cell.

##### `insertTable(arr, pos, options = {})`
- 从指定位置开始，插入一个表

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| arr | Array | Yes | - | 表格数据数组 |
| pos | Object \| String | Yes | - | 插入位置 |
| options | Object | No | {} | 配置选项 {align, border, width, height, background, color} |

##### `mergeCells(areas = null, mode = 'default')`
- 合并单元格，支持多种模式

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| areas | Array \| Object \| String \| null | No | null | 区域，默认取 activeAreas |
| mode | String | No | 'default' | 模式: 'default'\|'center'\|'content'\|'same' |

##### `moveArea(moveArea, targetArea): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| moveArea | Object \| string | Yes | - | - |
| targetArea | Object \| string | Yes | - | - |

**Returns**
- Type: `boolean`

##### `paddingArea(oArea = this.Canvas.lastPadding.oArea, targetArea = this.Canvas.lastPadding.targetArea, type = 'order')`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| oArea | Object \| string | No | this.Canvas.lastPadding.oArea | - |
| targetArea | Object \| string | No | this.Canvas.lastPadding.targetArea | - |
| type | string | No | 'order' | Fill area with series/copy/format. |

##### `paste(targetArea = null, options = {})`
- 粘贴数据到目标区域

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetArea | Object | No | null | 目标区域起始位置 { r, c } 或完整区域 |
| options | Object | No | {} | 粘贴选项 |
| options.mode | string | Yes | - | 粘贴模式 (PasteMode) |
| options.operation | string | Yes | - | 运算模式 (PasteOperation) |
| options.skipBlanks | boolean | Yes | - | 跳过空单元格 |
| options.transpose | boolean | Yes | - | 转置 |
| options.externalData | Object | Yes | - | 外部剪贴板数据（从系统剪贴板解析） |

##### `rangeSort(sortKeys, range?)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sortKeys | Array<{col:string \| number, order?:string, customOrder?:Array}> | Yes | - | - |
| range | RangeRef | No | - | - |

##### `rangeStrToNum(range): RangeNum`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | RangeStr | Yes | - | - |

**Returns**
- Type: `RangeNum`

##### `setBrush(keep = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| keep | boolean | No | false | - |

##### `setDataValidation(range, rule)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| rule | - | Yes | - | - |

##### `showAllHidCols()`

##### `showAllHidRows()`

##### `subtotal(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `textToColumns(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `ungroupCols(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `ungroupRows(range, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | - | Yes | - | - |
| options | Object | No | {} | - |

##### `unMergeCells(cellAd = null)`
- 解除合并，传入单元格地址或区域，默认取 activeAreas

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellAd | Object \| String \| null | No | null | 单元格地址或区域 |

##### `zoomIn(step = 0.1): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| step | number | No | 0.1 | - |

**Returns**
- Type: `number`

##### `zoomOut(step = 0.1): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| step | number | No | 0.1 | - |

**Returns**
- Type: `number`

##### `zoomToSelection(): number`

**Returns**
- Type: `number`
- Zoom to fit the selected area.

## Slicer/Slicer.js

### Slicer
- 切片器管理器<br>负责切片器的增删改查

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| map | Map<string, SlicerItem> | No | new Map() | - |
| sheet | Sheet | No | sheet | - |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| size | number | No | - |

#### Methods
##### `add(config): Slicer`
- 添加切片器

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | 配置对象 |

**Returns**
- Type: `Slicer`

##### `forEach(callback)`
- 遍历所有切片器

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

##### `get(id): Slicer | null`
- 获取切片器

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

**Returns**
- Type: `Slicer | null`

##### `getAll(): Array<Slicer>`
- 获取所有切片器

**Returns**
- Type: `Array<Slicer>`

##### `remove(id)`
- 删除切片器

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

## Sparkline/Sparkline.js

### Sparkline
- 迷你图管理器

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | 所属工作表 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| groups | Array<Object> | No | [] | 迷你图组列表 |
| map | Map<string, SparklineItem> | No | new Map() | 按位置索引的迷你图 Map： "R:C" -> SparklineItem |
| sheet | Sheet | No | sheet | 所属工作表 |

#### Methods
##### `add(config): SparklineItem | null`
- 添加迷你图

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | 配置对象 |
| config.cellRef | string \| Object | Yes | - | 单元格引用 "A1" 或 {r, c} |
| config.formula | string | Yes | - | 数据范围公式 |
| config.type | string | No | 'line' | 迷你图类型 |
| config.colors | Object | No | {} | 颜色配置 |

**Returns**
- Type: `SparklineItem | null`

##### `clearAllCache()`
- 清除所有缓存

##### `clearCache(cellRefOrRow, c?)`
- 清除指定迷你图的缓存

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRefOrRow | string \| Object \| number | Yes | - | 单元格引用 "A1"/{r,c} 或行索引 |
| c | number | No | - | 列索引（仅在第一个参数为行索引时使用） |

##### `get(cellRef): SparklineItem | null`
- 根据单元格位置获取迷你图

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string \| Object | Yes | - | 单元格引用 "A1" 或 {r, c} |

**Returns**
- Type: `SparklineItem | null`

##### `getAll(): Array<SparklineItem>`
- 获取所有迷你图实例

**Returns**
- Type: `Array<SparklineItem>`

##### `parse(xmlObj)`
- 从 xmlObj 解析迷你图数据

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | Sheet 的 XML 对象 |

##### `remove(cellRef): boolean`
- 删除迷你图

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string \| Object | Yes | - | 单元格引用 "A1" 或 {r, c} |

**Returns**
- Type: `boolean`

##### `toXmlObj(): Object | null`
- 导出为 XML 对象结构

**Returns**
- Type: `Object | null`

## Table/Table.js

### Table

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| size | - | No | 表格数量 |

#### Methods
##### `add(options = {}): TableItem`
- 添加/创建新表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

**Returns**
- Type: `TableItem`

##### `createFromSelection(area, options = {}): TableItem`
- 从选区创建表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | { s: {r, c}, e: {r, c} } |
| options | Object | No | {} | - |

**Returns**
- Type: `TableItem`

##### `forEach(callback)`
- 遍历所有表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

##### `fromJSON(jsonArray)`
- 从 JSON 恢复

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| jsonArray | Array | Yes | - | - |

##### `get(id): TableItem | null`
- 根据ID获取表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

**Returns**
- Type: `TableItem | null`

##### `getAll(): TableItem[]`

**Returns**
- Type: `TableItem[]`

##### `getByName(name): TableItem | null`
- 根据名称获取表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `TableItem | null`

##### `getTableAt(row, col): TableItem | null`
- 获取包含指定单元格的表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `TableItem | null`

##### `remove(id): boolean`
- 删除表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | 表格ID |

**Returns**
- Type: `boolean`

##### `removeByName(name): boolean`
- 根据名称删除表格

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `boolean`

##### `toJSON()`
- 转为 JSON 数组

## UndoRedo/UndoRedo.js

### UndoRedo

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | import('../Workbook/Workbook.js').default | Yes | - | - |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| maxStack | number | No | 50 | - |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| canRedo | boolean | No | - |
| canUndo | boolean | No | - |

#### Methods
##### `add(action): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| action | {undo: Function, redo: Function, activeCell?: {r:number,c:number}} | Yes | - | - |

**Returns**
- Add undo action into current transaction.

##### `begin(name = ''): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | No | '' | - |

**Returns**
- Begin a manual undo transaction.

##### `clear(): void`

**Returns**
- Clear undo/redo stacks and pending transaction.

##### `commit(): void`

**Returns**
- Commit current manual transaction.

##### `redo(): boolean`

**Returns**
- Type: `boolean`

##### `undo(): boolean`

**Returns**
- Type: `boolean`

## Utils/Utils.js

### Utils

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | SheetNext | Yes | - | - |

#### Methods
##### `cellNumToStr(cellNum): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellNum | {c:number,r:number} | Yes | - | - |

**Returns**
- Type: `string`

##### `cellStrToNum(cellStr): {c:number,r:number} | undefined`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellStr | string | Yes | - | - |

**Returns**
- Type: `{c:number,r:number} | undefined`

##### `charToNum(name): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `number`

##### `debounce(func, wait): Function`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| func | Function | Yes | - | - |
| wait | number | Yes | - | - |

**Returns**
- Type: `Function`
- Create a debounced function.

##### `decodeEntities(encodedString): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| encodedString | string | Yes | - | - |

**Returns**
- Type: `string`
- Decode HTML entities to plain text.

##### `async downloadImageToBase64(imageUrl): Promise<string | undefined>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| imageUrl | string | Yes | - | - |

**Returns**
- Type: `Promise<string | undefined>`
- Fetch image URL and return base64 data URL.

##### `emuToPx(emu): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| emu | number | Yes | - | - |

**Returns**
- Type: `number`
- Convert EMU to pixels.

##### `getFillFormula(formula, originalRow, originalCol, targetRow, targetCol): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formula | string | Yes | - | - |
| originalRow | number | Yes | - | - |
| originalCol | number | Yes | - | - |
| targetRow | number | Yes | - | - |
| targetCol | number | Yes | - | - |

**Returns**
- Type: `string`
- Build fill formula by shifting relative refs only.

##### `modal(options = {}): Promise<{bodyEl: HTMLElement} | null>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

**Returns**
- Type: `Promise<{bodyEl: HTMLElement} | null>`
- Open common modal panel.

##### `numToChar(num): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| num | number | Yes | - | - |

**Returns**
- Type: `string`

##### `objToArr(obj): Array<any>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | any | Yes | - | - |

**Returns**
- Type: `Array<any>`
- Normalize any value to array.

##### `pxToEmu(px): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| px | number | Yes | - | - |

**Returns**
- Type: `number`
- Convert pixels to EMU.

##### `rangeNumToStr(obj, absolute = false): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | {s:{r:number,c:number},e:{r:number,c:number}} | Yes | - | - |
| absolute | boolean | No | false | - |

**Returns**
- Type: `string`

##### `rgbToHex(rgb): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rgb | string | Yes | - | - |

**Returns**
- Type: `string`
- Convert rgb() string to hex color.

##### `sortObjectInPlace(obj, keyOrder): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | Object | Yes | - | - |
| keyOrder | Array<string> | Yes | - | - |

**Returns**
- Reorder object keys in-place.

##### `toast(message): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| message | any | Yes | - | - |

**Returns**
- Show toast message.

## Workbook/Workbook.js

### SheetNext

#### Constructor
- Create workbook instance.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dom | HTMLElement | Yes | - | Editor container element. |
| options | SheetNextInitOptions | No | {} | Initialization options. |
| options.licenseKey | string | No | - | License key. |
| options.locale | string | No | 'en-US' | Initial locale (e.g. 'en-US', 'zh-CN'). |
| options.locales | Object<string, Object> | No | - | Extra locale packs keyed by locale code. |
| options.menuRight | function | No | - | Callback `(defaultHTML: string) => string`. Receives the default right-menu HTML, return modified HTML. |
| options.menuList | function | No | - | Callback `(config: Array<{key: string, labelKey: string, groups: Array, contextual?: boolean}>) => Array`. Receives the default toolbar panel config array, return modified array. |
| options.AI_URL | string | No | - | AI relay endpoint URL. |
| options.AI_TOKEN | string | No | - | Optional bearer token for AI relay endpoint. |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| Action | Action | No | new Action(this) | - |
| AI | AI | No | new AI(this, options) | - |
| calcMode | 'auto' \| 'manual' | No | 'auto' | - |
| Canvas | Canvas | No | new Canvas(this) | - |
| containerDom | - | No | dom | - |
| DependencyGraph | DependencyGraph | No | new DependencyGraph(this) | - |
| Event | EventEmitter | No | new EventEmitter() | - |
| Formula | Formula | No | new Formula(this) | - |
| I18n | - | No | this._createI18n(options) | - |
| IO | IO | No | new IO(this) | - |
| Layout | Layout | No | new Layout(this, options) | - |
| License | License | No | new License(this, options.licenseKey) | - |
| namespace | string | No | this._setupGlobalNamespace() | - |
| Print | Print | No | new Print(this) | - |
| properties | Object | No | {} | - |
| sheets | Sheet[] | No | [] | - |
| UndoRedo | UndoRedo | No | new UndoRedo(this) | - |
| Utils | Utils | No | new Utils(this) | - |
| Xml | Xml | No | new Xml(this) | - |

#### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| activeSheet | Sheet | get/set | No | - |
| readOnly | boolean | get/set | No | - |
| workbookName | string | get/set | No | - |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| locale | string | No | - |

#### Methods
##### `addSheet(sheetName?): Sheet | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | No | - | - |

**Returns**
- Type: `Sheet | null`

##### `delSheet(name)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

##### `static getLocale(locale): Object | undefined`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

**Returns**
- Type: `Object | undefined`

##### `getSheet(sheetName): Sheet | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | Yes | - | - |

**Returns**
- Type: `Sheet | null`

##### `recalculate(currentSheetOnly = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| currentSheetOnly | boolean | No | false | - |

##### `static registerLocale(locale, messages): typeof SheetNext`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |
| messages | Object | Yes | - | - |

**Returns**
- Type: `typeof SheetNext`

##### `setLocale(locale)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

##### `t(key, params = {}): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| key | string | Yes | - | - |
| params | Record<string, any> | No | {} | - |

**Returns**
- Type: `string`

## Xml/Xml.js

### Xml
- Xml 解析与构建器

#### Constructor
**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | Yes | - | SheetNext 主实例 |
| obj | Object | No | - | XML 对象 |

#### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| obj | Object | No | obj ?? JSON.parse(JSON.stringify(snXml)) | XML 对象 |
| objTp | Object | No | snXml | XML 模板 |
| sharedStringsObj | Object | No | {} // 共享字符串去重对象清空 | 共享字符串去重对象 |
| SN | Object | No | SN | SheetNext 主实例 |
| Utils | Utils | No | SN.Utils | 工具方法集合 |

#### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| clrScheme | Object | No | 当前文件主题配色 |
| clrSchemeTp | Object | No | 主题模板配色 |
| override | Array | No | 覆盖关系 |
| relationship | Array | No | 关系集合 |
| sheets | Array | No | 所有工作表信息 |
| sst | Object | No | 共享字符串表 |
| styleSheet | Object | No | 样式表 |

#### Methods
##### `addSheet(sheetName, ops): void`
- 添加工作表

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | Yes | - | 工作表名称 |
| ops | Object | Yes | - | 初始化参数 |

##### `getRelsByTarget(targetPath, create = false): Object`
- 根据关系 target 路径获取 .rels 中的关系数组

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetPath | string | Yes | - | 目标路径 |
| create | boolean | No | false | 是否在不存在时创建 |

**Returns**
- Type: `Object`

##### `getWsRelsFileTarget(sheetRId, fileRId): string | null`
- 根据工作表依赖文件中的 RId 获取目标路径

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetRId | string | Yes | - | 工作表 RId |
| fileRId | string | Yes | - | 依赖文件 RId |

**Returns**
- Type: `string | null`

##### `resetStyle(): void`
- 重置样式表

##### `resolveHyperlinkById(rId): Object | null`
- 根据 RId 解析超链接

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rId | string | Yes | - | 关系 Id |

**Returns**
- Type: `Object | null`
