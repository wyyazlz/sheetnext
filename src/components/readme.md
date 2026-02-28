# Components

## Tooltip

```js
import { initTooltip } from './tooltip.js';

initTooltip(containerDom);
```

**HTML结构：**
```html
<button class="sn-tooltip" title="提示文字">按钮</button>
<!-- 或使用 data-title -->
<button class="sn-tooltip" data-title="提示文字" data-placement="bottom">按钮</button>
```

**属性：**
- `title` / `data-title` - 提示内容
- `data-placement` - 位置：`top`(默认) | `bottom` | `left` | `right`

---

## Dropdown

```js
import { initDropdown } from './dropdown.js';

initDropdown(containerDom);
```

**HTML结构：**
```html
<div class="sn-dropdown" data-keep-open="true">
    <button class="sn-dropdown-toggle sn-no-icon">菜单</button>
    <ul class="sn-dropdown-menu">
        <li>选项1</li>
        <li>
            带子菜单
            <ul class="sn-dropdown-submenu">
                <li>子选项</li>
            </ul>
        </li>
    </ul>
</div>
```

**属性：**
- `data-keep-open="true"` - 点击菜单项后保持打开
- `.sn-no-icon` - 加在 `sn-dropdown-toggle` 上，隐藏下拉小三角

**特性：**
- 自动边界检测，超出视口时翻转位置
- 支持多级子菜单
- 带关闭动画
