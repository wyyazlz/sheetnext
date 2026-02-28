/**
 * 事件管理模块
 * 负责为画布、滚动条、输入框等元素添加事件监听器
 */

import { canvasWheel, baseMousedown, canvasMousemove, canvasDblclick, rightDown, canvasTouchStart, canvasTouchMove, canvasTouchEnd } from "./InteractionHandler.js"
import { docKeyDown, docPaste, docCopy, docCut } from "./KeyboardHandler.js"
import { xDown, yDown, xTouchDown, yTouchDown } from "./scrollHandler.js"
import { formulaBarFocus, inputInput, formulaBarKeyDown } from "./inputHandler.js"
import { syncToolbarState } from "../Layout/ToolbarBuilder.js"

export function addEventListeners() {
    // 文档
    window.addEventListener('keydown', (event) => docKeyDown.call(this, event));
    document.addEventListener('paste', (event) => docPaste.call(this, event));
    document.addEventListener('copy', (event) => docCopy.call(this, event));
    document.addEventListener('cut', (event) => docCut.call(this, event));
    // 画布
    this.handleLayer.addEventListener('mousedown', (event) => baseMousedown.call(this, event));
    this.handleLayer.addEventListener('mousedown', (event) => rightDown.call(this, event));
    this.handleLayer.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    this.handleLayer.addEventListener('mousemove', (event) => canvasMousemove.call(this, event));
    this.handleLayer.addEventListener('wheel', (event) => canvasWheel.call(this, event));
    this.handleLayer.addEventListener('dblclick', (event) => canvasDblclick.call(this, event));
    // 触屏事件
    this.handleLayer.addEventListener('touchstart', (event) => canvasTouchStart.call(this, event), { passive: true });
    this.handleLayer.addEventListener('touchmove', (event) => canvasTouchMove.call(this, event), { passive: false });
    this.handleLayer.addEventListener('touchend', (event) => canvasTouchEnd.call(this, event));
    // 滚动条
    this.rollYS.addEventListener('mousedown', (event) => xDown.call(this, event));
    this.rollXS.addEventListener('mousedown', (event) => yDown.call(this, event));
    // 滚动条触摸事件（passive: false 允许阻止默认滚动）
    this.rollYS.addEventListener('touchstart', (event) => xTouchDown.call(this, event), { passive: false });
    this.rollXS.addEventListener('touchstart', (event) => yTouchDown.call(this, event), { passive: false });
    // 输入框
    this.input.addEventListener('input', () => inputInput.call(this))
    this.input.addEventListener('blur', () => {
        if (this.inputEditing) this.updInputValue();
    });
    this.formulaBar.addEventListener('focus', () => formulaBarFocus.call(this))
    this.formulaBar.addEventListener('keydown', (event => formulaBarKeyDown.call(this, event)))

    // 工具栏和菜单 tab mousedown 阻止 blur（保持编辑器焦点和选区）
    const toolsEl = this.SN.containerDom.querySelector('.sn-tools');
    const menuListEl = this.SN.containerDom.querySelector('.sn-menu-list');
    const preventBlur = (e) => { if (this.inputEditing) e.preventDefault(); };
    if (toolsEl) toolsEl.addEventListener('mousedown', preventBlur);
    if (menuListEl) menuListEl.addEventListener('mousedown', preventBlur);

    // selectionchange 同步工具栏状态（编辑时光标/选区变化更新工具栏）
    document.addEventListener('selectionchange', () => {
        if (!this.inputEditing || !this.currentCell) return;
        const { c, r } = this.currentCell;
        const cell = this.activeSheet.getCell(r, c);
        syncToolbarState(this.SN.containerDom, cell, this.SN.Layout?._toolbarBuilder?.menuConfig);
    });
}
