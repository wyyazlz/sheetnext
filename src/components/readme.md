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
- `.sn-menu-selectable` - 菜单项预留勾号位置；`.sn-selected` 显示勾号和主题色高亮

工具栏普通菜单支持 `selected: ctx => boolean`（或静态布尔值），`ctx` 为当前编辑器实例。例如纸张菜单配置 `selected: ctx => ctx.Print.ensureSettings()?.pageSetup.paperName === 'A4'`。打开菜单、执行菜单操作和刷新工具栏时均按实际设置更新选中态；执行命令或取消弹窗不会直接改变勾选。

通用下拉菜单在打开、计算位置之前，从菜单元素派发冒泡事件 `sn-dropdown-before-open`，用于懒加载内容和同步状态。

**特性：**
- 自动边界检测，超出视口时翻转位置
- 支持多级子菜单
- 带关闭动画
