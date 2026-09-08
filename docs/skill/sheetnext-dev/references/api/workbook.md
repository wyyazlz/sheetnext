# Workbook API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [SheetNext](#sheetnext)

## SheetNext
- Source: `src/core/Workbook/Workbook.js`

### Constructor
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
| options.menuList | function | No | - | Callback `(config: Array<{key: string, labelKey?: string, text?: string, groups?: Array, contextual?: boolean, contextType?: string, trigger?: 'panel'\|'action', action?: string, color?: string}>) => Array`. Receives the default toolbar panel config array, return modified array. Use `trigger: 'action'` to make a top tab run code directly without switching the toolbar panel. `color` customizes the top-tab text and accent color. |
| options.minimalToolbar | boolean | No | false | Start with the minimal toolbar enabled. |
| options.theme | Object \| string | No | - | UI theme overrides. Passing a string is treated as `primary`; object tokens include `primary`, `primaryHover`, `primaryActive`, `primarySoft`, `primarySoftHover`, `primaryBorder`, `primarySubtleBorder`, `primaryRing`, `primaryShadow`, `primaryLight`, and `primaryContrast`. |
| options.AI_URL | string | No | - | AI relay endpoint URL. |
| options.AI_TOKEN | string | No | - | Optional bearer token for AI relay endpoint. |

**Examples**
```js
import SheetNext from 'sheetnext';
import 'sheetnext/dist/sheetnext.css';

const container = document.querySelector('.sheetnext-container');
const SN = new SheetNext(container);
```
```js
const SN = new SheetNext(document.querySelector('.sheetnext-container'), {
  menuList(config) {
    return [
      ...config,
      {
        key: 'githubLink',
        text: 'GitHub',
        trigger: 'action',
        color: '#0969da',
        action: "window.open('https://github.com/wyyazlz/sheetnext', '_blank', 'noopener,noreferrer')"
      }
    ];
  }
});
```

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| containerDom | HTMLElement | No | dom | - |
| namespace | string | No | this._setupGlobalNamespace() | - |
| calcMode | 'auto' \| 'manual' | No | 'auto' | - |
| sheets | Sheet[] | No | [] | - |
| properties | Object | No | {} | - |
| Event | EventEmitter | No | new EventEmitter() | - |
| Utils | Utils | No | new Utils(this) | - |
| I18n | I18n | No | this._createI18n(options) | - |
| Print | Print | No | new Print(this) | - |
| Layout | Layout | No | new Layout(this, options) | - |
| AI | AI | No | new AI(this, options, this.#license) | - |
| IO | IO | No | new IO(this, this.#license) | - |
| Formula | Formula | No | new Formula(this) | - |
| UndoRedo | UndoRedo | No | new UndoRedo(this) | - |
| Canvas | Canvas | No | new Canvas(this) | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| activeSheet | Sheet | get/set | No | - |
| workbookName | string | get/set | No | - |
| readOnly | boolean | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| locale | string | No | - |

### Methods
#### `static registerLocale(locale, messages): typeof SheetNext`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |
| messages | Object | Yes | - | - |

**Returns**
- Type: `typeof SheetNext`

**Examples**
```js
import SheetNext from 'sheetnext';
import zhCN from 'sheetnext/locales/zh-CN';

SheetNext.registerLocale('zh-CN', zhCN);
```

#### `static getLocale(locale): Object | undefined`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

**Returns**
- Type: `Object | undefined`

#### `setTheme(theme = {}): Object`
- Apply UI theme colors to this workbook.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| theme | Object \| string | No | {} | Theme overrides. A string is treated as the primary color. |

**Returns**
- Type: `Object`

#### `addSheet(sheetName?): Sheet | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | No | - | - |

**Returns**
- Type: `Sheet | null`

**Examples**
```js
const budgetSheet = SN.addSheet('Budget 2026');

if (!budgetSheet) throw new Error('Could not create sheet');
budgetSheet.getCell('A1').editVal = 'Month';
budgetSheet.getCell('B1').editVal = 'Amount';
```

#### `delSheet(name): boolean | void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `boolean | void`

#### `getSheet(sheetName): Sheet | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | Yes | - | - |

**Returns**
- Type: `Sheet | null`

**Examples**
```js
const sheet = SN.getSheet('Budget 2026');
if (sheet) {
  sheet.getCell('B2').editVal = 128000;
}
```

#### `moveSheet(sheetOrName, targetIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetOrName | Sheet \| string | Yes | - | - |
| targetIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `recalculate(currentSheetOnly = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| currentSheetOnly | boolean | No | false | - |

**Examples**
```js
SN.recalculate();
SN.recalculate(true);
```

#### `setLocale(locale)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

**Examples**
```js
SN.setLocale('zh-CN');
```

#### `t(key, params = {}): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| key | string | Yes | - | - |
| params | Record<string, any> | No | {} | - |

**Returns**
- Type: `string`
