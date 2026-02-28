
/**
 * 通过单元格索引获取透视表（基于 location.ref 判断）
 */
function getPivotTableByCellIndex(sheet, row, col) {
    if (!sheet.PivotTable || sheet.PivotTable.size === 0) return null;

    for (const pt of sheet.PivotTable.getAll()) {
        const ref = pt.location?.ref;
        if (!ref) continue;

        // 解析区域引用 (如 "A1:C18")
        const parts = ref.split(':');
        const start = parseCellRef(parts[0]);
        const end = parts.length > 1 ? parseCellRef(parts[1]) : start;
        if (!start || !end) continue;

        if (row >= start.r && row <= end.r && col >= start.c && col <= end.c) {
            return pt;
        }
    }
    return null;
}

/**
 * 解析单元格引用 (如 "A1" -> {r: 0, c: 0})
 */
function parseCellRef(cellRef) {
    const match = cellRef.replace(/\$/g, '').match(/^([A-Z]+)(\d+)$/i);
    if (!match) return null;
    const colStr = match[1].toUpperCase();
    const row = parseInt(match[2], 10) - 1;
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    return { r: row, c: col - 1 };
}

const downBeforeFuns = [] // 鼠标点下后执行一次后移除，用于取消上次选中和悬浮框
const downFuns = {
    // 行列调整
    'row-col-resize': function (event, info) {

        const snDragLine = this.lazyDOM.dragLine;

        const { mouseX: x, mouseY: y, sheet } = info;
        const layer = this.handleLayer;

        let cellIndex = this.getCellIndex(event);
        const cellInfo = sheet.getCellInViewInfo(cellIndex.r, cellIndex.c);

        // 边界检测优化
        const isColResize = y < sheet.headHeight;
        const isRowResize = x < sheet.indexWidth;
        const nearBoundary = 5;

        // 调整到前一个单元格
        if (isRowResize && Math.abs(cellInfo.y - y) <= nearBoundary && cellIndex.r > 0) {
            cellIndex.r = sheet._findBeforeShowRow(cellIndex.r, 1);
        } else if (isColResize && Math.abs(cellInfo.x - x) <= nearBoundary && cellIndex.c > 0) {
            cellIndex.c = sheet._findBeforeShowCol(cellIndex.c, 1);
        }

        // 获取当前尺寸
        const isHorizontal = layer.style.cursor === 'ew-resize';
        const startSize = isHorizontal ? sheet.getCol(cellIndex.c).width : sheet.getRow(cellIndex.r).height;

        let readyExe = null
        const mouseMoveHandler = (event) => {
            const { x: moveX, y: moveY } = this.getEventPosition(event);
            const offset = isHorizontal ? moveX - x : moveY - y;
            const newSize = Math.max(offset + startSize, 3);
            // 统一设置拖拽线样式
            Object.assign(snDragLine.style, {
                display: 'block',
                cursor: layer.style.cursor,
                width: isHorizontal ? '1px' : `${this.width}px`,
                height: isHorizontal ? `${this.height}px` : '1px',
                transform: isHorizontal ? `translateX(${event.offsetX}px)` : `translateY(${event.offsetY}px)`
            });
            // 准备执行的操作
            readyExe = () => {
                if (isHorizontal) {
                    sheet.getCol(cellIndex.c).width = newSize;
                } else {
                    sheet.getRow(cellIndex.r).height = newSize;
                }
            };
        };

        layer.addEventListener('mousemove', mouseMoveHandler);
        window.addEventListener('mouseup', () => {
            layer.removeEventListener('mousemove', mouseMoveHandler);
            readyExe?.();
            readyExe = null;
            layer.style.cursor = 'default';
            snDragLine.style.display = 'none';
            this.disableMoveCheck = false;
            this.r();
        }, { once: true });

    },
    // 图纸移动
    'drawing': function (event, info) {
        const { mouseX, mouseY, sheet } = info;
        let d = this.activeDrawing
        d.active = true
        let offsetX = mouseX - d.position.x
        let offsetY = mouseY - d.position.y;
        const layer = this.handleLayer;
        const mouseMoveHandler = (e) => {
            const { x: mx, y: my } = this.getEventPosition(e);
            if (d?.position) {
                d.position.x = mx - offsetX;
                d.position.y = my - offsetY;
                if (d.position.x < sheet.indexWidth) d.position.x = sheet.indexWidth;
                if (d.position.y < sheet.headHeight) d.position.y = sheet.headHeight;
            }
            d.moving = true
            this.r('s');
        };
        layer.addEventListener('mousemove', mouseMoveHandler);
        window.addEventListener('mouseup', () => {
            layer.removeEventListener('mousemove', mouseMoveHandler);
            // 只有在真正发生过移动时，才更新 area
            if (d.moving) {
                const { x, y, w, h } = d.position;
                d.startCell = this.getCellIndexByPosition(x, y);
                const cellInfo = sheet.getCellInViewInfo(d.startCell.r, d.startCell.c, false);
                d.offsetX = x - cellInfo.x;
                d.offsetY = y - cellInfo.y;
                d.width = w
                d.height = h
                sheet.getCell(d.area.e.r, d.area.e.c)
            }
            d.moving = false;
            this.disableMoveCheck = false
            downBeforeFuns.push(() => d.active = false);
            this.r()
        }, { once: true });
        this.r('s');
    },
    // 活动区域移动
    'active-border': function (event, info) {
        const layer = this.handleLayer;
        const { sheet } = info
        const a = this.activeSheet.activeAreas
        const mouseMoveHandler = (event2) => {
            const newIndex = this.getCellIndex(event2);
            const newArea = updateRegionBounds.call(this, a, this.getCellIndex(event), newIndex); // 获取新区域
            if (Object.keys(this.mpBorder).length < 1 || newArea.s.r != this.mpBorder.s.r || newArea.s.c != this.mpBorder.s.c) {
                this.mpBorder = newArea;
                this.r();
            }
        };
        layer.addEventListener('mousemove', mouseMoveHandler);
        layer.addEventListener('mouseup', () => {
            layer.removeEventListener('mousemove', mouseMoveHandler);
            if (Object.keys(this.mpBorder).length > 0) {
                const moved = sheet.moveArea(a[a.length - 1], this.mpBorder) // 更新表数据
                if (moved !== false) {
                    sheet.activeAreas = [this.mpBorder]
                }
            }
            this.mpBorder = {}
            this.disableMoveCheck = false
            this.r();
        }, { once: true });
    },
    // 填充柄
    'fill-handle': function (event, info) {
        const layer = this.handleLayer;
        const sheet = this.activeSheet;
        const a = this.activeSheet.activeAreas
        const mouseMoveHandler = (event) => {
            const newArea = getPaddingArea.call(this, this.getCellIndex(event));
            if (!newArea) {
                if (Object.keys(this.mpBorder).length > 0) {
                    this.mpBorder = {}
                    this.r();
                }
                return
            }
            if (Object.keys(this.mpBorder).length < 1 || newArea.s.r != this.mpBorder.s.r || newArea.s.c != this.mpBorder.s.c || newArea.e.r != this.mpBorder.e.r || newArea.e.c != this.mpBorder.e.c) {
                this.mpBorder = newArea;
                this.r();
            }
        }
        layer.addEventListener('mousemove', mouseMoveHandler);
        layer.addEventListener('mouseup', () => {
            layer.removeEventListener('mousemove', mouseMoveHandler);
            if (Object.keys(this.mpBorder).length > 0) {
                this.lastPadding = { oArea: a[a.length - 1], targetArea: this.mpBorder }
                const padded = sheet.paddingArea();
                if (padded === false) {
                    this.mpBorder = {}
                    this.r();
                    this.disableMoveCheck = false
                    return;
                }
                sheet.activeAreas = [mergeExcelRanges.call(this, this.mpBorder, a[a.length - 1])]

                // 触发 Volatile 函数重算
                if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                    this.SN.DependencyGraph.recalculateVolatile();
                }

                // 显示 paddingType 按钮
                const lastArea = sheet.activeAreas[sheet.activeAreas.length - 1];
                const areaInfo = sheet.getAreaInviewInfo(lastArea);
                if (areaInfo.inView) {
                    this.paddingType.style.left = `${this.toPhysical(areaInfo.x + areaInfo.w) + 5}px`;
                    this.paddingType.style.top = `${this.toPhysical(areaInfo.y + areaInfo.h) + 5}px`;
                    this.paddingType.style.display = 'block';
                    downBeforeFuns.push(() => this.paddingType.style.display = 'none');
                }

                this.mpBorder = {}
            }
            this.r();
            this.disableMoveCheck = false
        }, { once: true });
    },
    // 图纸8个点位拖拉
    'd-left-top': function (e, i) { dHandleResize.call(this, e, i, 'left-top'); },
    'd-right-top': function (e, i) { dHandleResize.call(this, e, i, 'right-top'); },
    'd-left-bottom': function (e, i) { dHandleResize.call(this, e, i, 'left-bottom'); },
    'd-right-bottom': function (e, i) { dHandleResize.call(this, e, i, 'right-bottom'); },
    'd-left': function (e, i) { dHandleResize.call(this, e, i, 'left'); },
    'd-right': function (e, i) { dHandleResize.call(this, e, i, 'right'); },
    'd-top': function (e, i) { dHandleResize.call(this, e, i, 'top'); },
    'd-bottom': function (e, i) { dHandleResize.call(this, e, i, 'bottom'); },
    // 图纸旋转
    'd-rotate': function (event, info) {
        const d = this.activeDrawing;
        d.active = true;
        const pos = d.position;
        const cx = pos.x + pos.w / 2;
        const cy = pos.y + pos.h / 2;

        const { x: mx0, y: my0 } = this.getEventPosition(event);
        // 从12点方向计算起始角度（atan2(dx, -dy) 使得正上方为0，顺时针为正）
        const startAngle = Math.atan2(mx0 - cx, -(my0 - cy));
        const startRotation = d._rotation || 0;

        const layer = this.handleLayer;
        layer.style.cursor = 'grabbing';

        const mouseMoveHandler = (e) => {
            const { x: mx, y: my } = this.getEventPosition(e);
            const currentAngle = Math.atan2(mx - cx, -(my - cy));
            let deg = startRotation + (currentAngle - startAngle) * 180 / Math.PI;
            // Shift键按住时吸附到15度整数倍
            if (e.shiftKey) deg = Math.round(deg / 15) * 15;
            d.rotation = deg;
            this.r('s');
        };

        layer.addEventListener('mousemove', mouseMoveHandler);
        window.addEventListener('mouseup', () => {
            layer.removeEventListener('mousemove', mouseMoveHandler);
            layer.style.cursor = 'default';
            this.disableMoveCheck = false;
            downBeforeFuns.push(() => d.active = false);
            this.r('s');
        }, { once: true });
        this.r('s');
    },
}

// 这里实时检查当前鼠标移动到哪个区域了，并保存到：moveAreaName，当鼠标按下时判断moveAreaName进行对应的事件，并停止检查。
// 检查顺序从性能高到低，这样中途可retrun
let lastMoveTime = 0;
export function canvasMousemove(event) {
    const now = performance.now();
    if (now - lastMoveTime < 20) return; // 节流
    lastMoveTime = now;

    this.moveAreaName = null;
    if (this.disableMoveCheck) return;
    const sheet = this.activeSheet;
    if (sheet._brush.area) return this.handleLayer.style.cursor = 'cell';

    const { x, y } = this.getEventPosition(event);

    const moveRowCol = getMouseInRC.call(this, x, y, event, sheet);
    if (moveRowCol) return this.handleLayer.style.cursor = moveRowCol;

    const moveBorder = getMouseInBorder.call(this, x, y);
    if (moveBorder) return this.handleLayer.style.cursor = moveBorder;

    const moveDrawing = getMouseInDrawing.call(this, x, y);
    if (moveDrawing) return this.handleLayer.style.cursor = moveDrawing;

    // 检查是否hover在批注单元格上
    const cellIndex = this.getCellIndex(event);
    const cellRef = sheet.SN.Utils.cellNumToStr(cellIndex);
    const comment = sheet.Comment.get(cellRef);

    // 检查超链接tooltip
    const hoverCell = sheet.getCell(cellIndex.r, cellIndex.c);
    if (hoverCell.hyperlink?.tooltip) {
        const { x: cx, y: cy, w, h } = sheet.getCellInViewInfo(cellIndex.r, cellIndex.c);
        this.hyperlinkTip.textContent = hoverCell.hyperlink.tooltip;
        this.hyperlinkTip.style.left = `${this.toPhysical(cx)}px`;
        this.hyperlinkTip.style.top = `${this.toPhysical(cy + h) + 2}px`;
        this.hyperlinkTip.style.display = 'block';
    } else {
        this.hyperlinkTip.style.display = 'none';
    }

    if (comment) {
        // hover在有批注的单元格上
        if (this.hoveredComment !== comment) {
            this.hoveredComment = comment;
            this.r();
        }
    } else if (this.hoveredComment) {
        // 离开批注单元格，检查当前选中单元格是否有批注
        const activeCell = sheet.activeCell;
        if (activeCell) {
            const selRef = sheet.SN.Utils.cellNumToStr(activeCell);
            const selComment = sheet.Comment.get(selRef);
            // 如果选中单元格的批注就是当前显示的批注，保持显示
            if (selComment === this.hoveredComment) {
                this.handleLayer.style.cursor = 'default';
                return;
            }
        }
        // 否则关闭批注框
        this.hoveredComment = null;
        this.r();
    }

    this.handleLayer.style.cursor = 'default';
}

export function canvasDblclick(event) {
    const { x, y } = this.getEventPosition(event);
    const drawing = getDrawingByPoint.call(this, x, y);
    if (drawing && drawing.type === 'shape' && drawing.isTextBox && !drawing.isConnector) {
        if (this.inputEditing) this.updInputValue();
        this.activeDrawing = drawing;
        drawing.active = true;
        this.r('s');
        startTextBoxInlineEditing.call(this, drawing);
        return;
    }
    finishTextBoxInlineEditing.call(this, true);
    this.showCellInput(false, true);
}

export function baseMousedown(event) {
    finishTextBoxInlineEditing.call(this, true);
    setTimeout(() => this.input.focus());
    downBeforeFuns.splice(0).forEach(fun => fun());
    if (event.button == 2) return
    const { left, top, width, height } = this.handleLayer.getBoundingClientRect();
    const sheet = this.activeSheet
    const mouseX = this.toLogical(event.clientX - left);
    const mouseY = this.toLogical(event.clientY - top);

    this.disableMoveCheck = true
    // 需要拖拉的事件汇总
    if (this.moveAreaName) return downFuns[this.moveAreaName].call(this, event, { left, top, mouseX, mouseY, sheet });

    this.disableMoveCheck = false
    const start = this.getCellIndex(event); // 获取点击时的单元格索引
    let downCell = sheet.getCell(start.r, start.c); // 按下单元格索引
    if (downCell.master) downCell = sheet.getCell(downCell.master.r, downCell.master.c)

    const toggleRow = downCell?.row?.rIndex ?? start.r;
    const toggleCol = downCell?.cIndex ?? start.c;
    const pivotTableForToggle = getPivotTableByCellIndex(sheet, toggleRow, toggleCol);
    if (pivotTableForToggle) {
        const toggleInfo = pivotTableForToggle.getRowToggleInfo(toggleRow, toggleCol);
        const rect = downCell?._pivotToggleRect;
        if (toggleInfo?.hasToggle && rect &&
            mouseX >= rect.x && mouseX <= rect.x + rect.w &&
            mouseY >= rect.y && mouseY <= rect.y + rect.h) {
            const cell = sheet.getCell(start.r, start.c);
            sheet.activeCell = { c: cell.master?.c ?? start.c, r: cell.master?.r ?? start.r };
            if (event.ctrlKey) {
                sheet.activeAreas.push({ s: { c: start.c, r: start.r }, e: { c: start.c, r: start.r } });
            } else {
                sheet.activeAreas = [];
            }
            pivotTableForToggle.toggleRowCollapsed(toggleInfo.level, toggleInfo.values);
            pivotTableForToggle.render();
            return;
        }
    }

    // #region 按下的是超链接需要显示跳转按钮
    if (downCell.hyperlink) {
        const { x, y, w, h } = sheet.getCellInViewInfo(start.r, start.c);
        this.hyperlinkJumpBtn.style.top = `${this.toPhysical(y) - 2}px`
        this.hyperlinkJumpBtn.style.left = `${this.toPhysical(w + x) + 5}px`
        this.hyperlinkJumpBtn.style.display = `block`
        downBeforeFuns.push(() => this.hyperlinkJumpBtn.style.display = `none`)
    }
    // #endregion

    // #region 按下的的单元格需要数据验证提示
    if (downCell?._dataValidation) {
        const { x, y, w, h } = sheet.getCellInViewInfo(start.r, start.c);
        if (downCell._dataValidation.showInputMessage && downCell._dataValidation.prompt) { // 提示框
            this.validTip.querySelector('div').innerText = downCell._dataValidation.promptTitle ?? ""
            this.validTip.querySelector('span').innerText = downCell._dataValidation.prompt ?? ""
            this.validTip.style.top = `${this.toPhysical(y + h) + 5}px`
            this.validTip.style.left = `${this.toPhysical(x)}px`
            this.validTip.style.display = `block`
            downBeforeFuns.push(() => this.validTip.style.display = `none`)
        }
        if (downCell._dataValidation?.type == 'list' && downCell?._dataValidation?.showDropDown) {
            this.validSelect.style.left = `${this.toPhysical(w + x) + 4}px`
            this.validSelect.style.top = `${this.toPhysical(y + h) - 20}px`
            this.validSelect.style.display = `block`
            const ns = this.SN.namespace; // 获取真实的命名空间
            const list = downCell?._dataValidation.formula1.map(item => `<li onclick="${ns}.activeSheet.getCell(${downCell.row.rIndex},${downCell.cIndex}).editVal='${item}';${ns}._r();">${item}</li>`).join("")
            this.validSelect.querySelector("ul").innerHTML = list
            downBeforeFuns.push(() => this.validSelect.style.display = `none`)
        }
    }
    // #endregion

    // #region 检查是否点击了筛选图标
    const filterHit = this.getFilterIconAtPosition(mouseX, mouseY);
    if (filterHit) {
        const { colIndex: filterColIndex, scopeId } = filterHit;
        const autoFilter = sheet.AutoFilter;
        if (autoFilter?.isScopeEnabled(scopeId)) {
            const headerRow = autoFilter.getScopeHeaderRow(scopeId);
            // 计算面板显示位置
            const colInfo = sheet.getCellInViewInfo(headerRow, filterColIndex);
            const panelX = colInfo.x + colInfo.w;
            const panelY = colInfo.y + colInfo.h;
            // 显示筛选面板
            this.filterPanelManager.show(filterColIndex, this.toPhysical(panelX), this.toPhysical(panelY), scopeId);
            return;
        }
    }
    // #endregion

    // #region 检查是否点击了透视表行标签筛选
    const pivotFilterRect = downCell?._pivotFilterRect;
    if (pivotFilterRect &&
        mouseX >= pivotFilterRect.x && mouseX <= pivotFilterRect.x + pivotFilterRect.w &&
        mouseY >= pivotFilterRect.y && mouseY <= pivotFilterRect.y + pivotFilterRect.h) {
        const pivotTableForFilter = getPivotTableByCellIndex(sheet, toggleRow, toggleCol);
        if (pivotTableForFilter && downCell?._pivotFilter) {
            const panelX = pivotFilterRect.x + pivotFilterRect.w;
            const panelY = pivotFilterRect.y + pivotFilterRect.h;
            this.filterPanelManager?.showPivotFilter?.(
                pivotTableForFilter,
                downCell._pivotFilter,
                this.toPhysical(panelX),
                this.toPhysical(panelY)
            );
            return;
        }
    }
    // #endregion

    // #region 按下的单元格或行列标,需要选中单元格或行列
    const p = sheet.protection;
    const canSelectLocked = p.enabled ? p.selectLockedCells : true;
    const canSelectUnlocked = p.enabled ? p.selectUnlockedCells : true;
    let lastSelectToastAt = 0;
    const toastSelectOnce = (msg) => {
        const now = Date.now();
        if (now - lastSelectToastAt < 500) return;
        lastSelectToastAt = now;
        sheet.SN?.Utils?.toast?.(msg);
    };
    const getSelectionBlockReason = (area) => {
        if (!p.enabled) return '';
        if (!canSelectLocked && !canSelectUnlocked) return sheet.SN.t('core.sheet.sheetProtection.message.selectionBlocked');
        if (!area) return '';
        const range = typeof area === 'string' ? sheet.rangeStrToNum(area) : area;
        if (!range?.s || !range?.e) return '';
        const sr = Math.min(range.s.r, range.e.r);
        const er = Math.max(range.s.r, range.e.r);
        const sc = Math.min(range.s.c, range.e.c);
        const ec = Math.max(range.s.c, range.e.c);
        for (let r = sr; r <= er; r++) {
            for (let c = sc; c <= ec; c++) {
                const cell = sheet.getCell(r, c);
                const locked = p.isCellLocked(cell);
                if (locked && !canSelectLocked) return sheet.SN.t('core.sheet.sheetProtection.message.cannotSelectLocked');
                if (!locked && !canSelectUnlocked) return sheet.SN.t('core.sheet.sheetProtection.message.cannotSelectUnlocked');
            }
        }
        return '';
    };
    const blockSelection = (area) => {
        const reason = getSelectionBlockReason(area);
        if (!reason) return false;
        toastSelectOnce(reason);
        return true;
    };
    if (mouseY < sheet.headHeight && mouseX < sheet.indexWidth) { // 全选
        const a = { s: { c: 0, r: 0 }, e: { c: sheet.colCount - 1, r: sheet.rowCount - 1 } };
        if (blockSelection(a)) return;
        sheet.activeCell = { c: 0, r: 0 }
        sheet.activeAreas = [a];
        return this.r('s')
    } else if (mouseY < sheet.headHeight) { // 选择一整列
        const a = { s: { c: start.c, r: 0 }, e: { c: start.c, r: sheet.rowCount - 1 } }
        if (blockSelection(a)) return;
        sheet.activeCell = { c: start.c, r: 0 }
        event.ctrlKey ? sheet.activeAreas.push(a) : sheet.activeAreas = [a];
    } else if (mouseX < sheet.indexWidth) { // 选择一整行
        const a = { s: { c: 0, r: start.r }, e: { c: sheet.colCount - 1, r: start.r } }
        if (blockSelection(a)) return;
        sheet.activeCell = { c: 0, r: start.r }
        event.ctrlKey ? sheet.activeAreas.push(a) : sheet.activeAreas = [a]
    } else {
        const cell = sheet.getCell(start.r, start.c); // 可能是合并的单元格
        const singleArea = sheet._getExpandMaxArea({ s: { c: start.c, r: start.r }, e: { c: start.c, r: start.r } });
        if (blockSelection(singleArea)) return;
        sheet.activeCell = { c: cell.master?.c ?? start.c, r: cell.master?.r ?? start.r }
        if (event.ctrlKey) { // 按下ctrl
            sheet.activeAreas.push({ s: { c: start.c, r: start.r }, e: { c: start.c, r: start.r } })
        } else {
            sheet.activeAreas = [];
        }
    }
    if (!this.inputEditing) this.r('s')
    // #endregion

    // 检查是否点击了透视表区域（使用单元格索引判断，更可靠）
    this.SN.Layout?.closePivotPanel?.();

    areaSel.call(this, start, left, top, width, height) // 区域选择
}

export function rightDown(event) {
    if (event.button != 2) return
    const a = this.activeSheet.activeAreas
    const rect = this.handleLayer.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const logicalX = this.toLogical(x);
    const logicalY = this.toLogical(y);
    const cellIndex = this.getCellIndex(event);
    const side = this.activeSheet.vi.side
    if (logicalX < this.activeSheet.indexWidth) { // 列索引右键
        if (a[a.length - 1].s.c != side.l || a[a.length - 1].e.c != side.r || cellIndex.r < a[a.length - 1].s.r || cellIndex.r > a[a.length - 1].e.r) {
            this.activeSheet.activeAreas = [{ s: { c: side.l, r: cellIndex.r }, e: { c: side.r, r: cellIndex.r } }]
            this.r('s')
        }
        const rowRC = this.lazyDOM.rowContextMenu;
        rowRC.style.display = "block";
        rowRC.style.left = x + 5 + 'px'
        rowRC.style.top = y + 5 + 'px'
        downBeforeFuns.push(() => rowRC.style.display = "none");
    } else if (logicalY < this.activeSheet.headHeight) { // 列头右键
        if (a[a.length - 1].s.r != side.t || a[a.length - 1].e.r != side.b || cellIndex.c < a[a.length - 1].s.c || cellIndex.c > a[a.length - 1].e.c) {
            this.activeSheet.activeAreas = [{ s: { c: cellIndex.c, r: side.t }, e: { c: cellIndex.c, r: side.b } }]
            this.r('s')
        }
        const colRC = this.lazyDOM.colContextMenu;
        colRC.style.display = "block";
        colRC.style.left = x + 5 + 'px'
        colRC.style.top = y + 5 + 'px'
        downBeforeFuns.push(() => colRC.style.display = "none")
    } else if (this.moveAreaName == 'drawing') {
        const d = this.activeDrawing
        d.active = true
        downBeforeFuns.push(() => d.active = false)
        this.r('s')
        const drawingRC = this.lazyDOM.drawingContextMenu;
        drawingRC.style.display = "block";
        drawingRC.style.left = x + 5 + 'px'
        drawingRC.style.top = y + 5 + 'px'
        downBeforeFuns.push(() => drawingRC.style.display = "none")
    } else {
        if (cellIndex.r < a[a.length - 1].s.r || cellIndex.r > a[a.length - 1].e.r || cellIndex.c < a[a.length - 1].s.c || cellIndex.c > a[a.length - 1].e.c) {
            this.activeSheet.activeCell = cellIndex
            this.activeSheet.activeAreas = [];
            this.r('s')
        }
        const cellRC = this.lazyDOM.cellContextMenu;
        cellRC.style.display = "block";
        cellRC.style.left = x + 5 + 'px'
        cellRC.style.top = y + 5 + 'px'
        downBeforeFuns.push(() => cellRC.style.display = "none")
    }
}

// 画布监听
export function canvasWheel(event) {
    // 如果禁用滚动，则直接返回
    if (this.disableScroll) return;

    // Ctrl/⌘ + 滚轮缩放
    if (event.ctrlKey || event.metaKey) {
        event.preventDefault();
        const current = this.activeSheet?.zoom ?? 1;
        const delta = event.deltaY < 0 ? 0.1 : -0.1;
        const next = Math.round((current + delta) * 100) / 100;
        if (this.activeSheet) this.activeSheet.zoom = next;
        return;
    }

    const sheet = this.activeSheet;
    const frozenRows = sheet._frozenRows || 0;
    const frozenCols = sheet._frozenCols || 0;

    if (!sheet.vi) return;

    const maxRow = sheet.vi.side?.b ?? (sheet.rowCount - 1);
    const maxCol = sheet.vi.side?.r ?? (sheet.colCount - 1);

    // 更新滚动位置（考虑冻结限制）
    if (event.deltaY > 0) {
        const cachedNextRow = sheet.vi.rowArr?.[3];
        const newRow = Number.isFinite(cachedNextRow)
            ? cachedNextRow
            : sheet._findNextShowRow(sheet.vi.rowArr[0], 3);
        sheet.vi.rowArr[0] = Math.min(Math.max(newRow, frozenRows), maxRow);
    } else if (event.deltaY < 0) {
        const newRow = sheet._findBeforeShowRow(sheet.vi.rowArr[0], 3);
        sheet.vi.rowArr[0] = Math.max(newRow, frozenRows); // 不能小于冻结行数
    } else if (event.deltaX > 0) {
        const cachedNextCol = sheet.vi.colArr?.[3];
        const newCol = Number.isFinite(cachedNextCol)
            ? cachedNextCol
            : sheet._findNextShowCol(sheet.vi.colArr[0], 3);
        sheet.vi.colArr[0] = Math.min(Math.max(newCol, frozenCols), maxCol);
    } else if (event.deltaX < 0) {
        const newCol = sheet._findBeforeShowCol(sheet.vi.colArr[0], 3);
        sheet.vi.colArr[0] = Math.max(newCol, frozenCols); // 不能小于冻结列数
    } else return
    // 重新渲染
    this.r();
    // 更新滚动条位置
    this.updateScrollBar();
    // 阻止默认滚动行为
    event.preventDefault();
    if (this.inputEditing) {
        const blurEvent = new Event('blur');
        this.input.dispatchEvent(blurEvent);
    }
}

// 触屏滑动 - 触摸开始
export function canvasTouchStart(event) {
    // 如果禁用滚动，则直接返回
    if (this.disableScroll) return;

    // 只处理单指触控
    if (event.touches.length !== 1) return;

    // 如果当前有其他操作（如拖拽调整大小），不处理滑动
    if (this.disableMoveCheck) return;

    const touch = event.touches[0];
    const rect = this.handleLayer.getBoundingClientRect();
    const x = touch.clientX - rect.left;
    const y = touch.clientY - rect.top;

    this.touchStartInfo = {
        x: touch.clientX,
        y: touch.clientY,
        localX: x,
        localY: y,
        time: Date.now(),
        scrollX: this.activeSheet.vi.colArr[0],
        scrollY: this.activeSheet.vi.rowArr[0],
        moved: false,
        lastRenderTime: Date.now() // 记录上次渲染时间
    };

    // 清除之前的CSS变换
    if (this.canvasContainer) {
        this.canvasContainer.style.transform = '';
    }
}

// 触屏滑动 - 触摸移动
export function canvasTouchMove(event) {
    // 如果禁用滚动，则直接返回
    if (this.disableScroll) return;

    if (!this.touchStartInfo || event.touches.length !== 1) return;

    const touch = event.touches[0];
    const deltaX = touch.clientX - this.touchStartInfo.x;
    const deltaY = touch.clientY - this.touchStartInfo.y;
    const logicalDeltaX = deltaX / (this.activeSheet?.zoom ?? 1);
    const logicalDeltaY = deltaY / (this.activeSheet?.zoom ?? 1);

    // 检测是否开始移动（移动超过5px才认为是滑动）
    if (!this.touchStartInfo.moved && (Math.abs(deltaX) > 5 || Math.abs(deltaY) > 5)) {
        this.touchStartInfo.moved = true;
    }

    if (!this.touchStartInfo.moved) return;

    // 阻止默认滚动行为
    event.preventDefault();

    const sheet = this.activeSheet;
    const now = Date.now();

    // 1. 立即应用CSS变换，提供视觉流畅感（无渲染开销）
    if (this.canvasContainer) {
        this.canvasContainer.style.transform = `translate(${deltaX}px, ${deltaY}px)`;
    }

    // 2. 每隔150ms进行一次真实渲染（降低渲染频率）
    const renderInterval = 150;
    if (now - this.touchStartInfo.lastRenderTime < renderInterval) {
        return; // 跳过本次渲染
    }

    // 根据滑动距离计算需要移动的行/列数
    const threshold = 15;  // 缩小阈值提高响应频率
    const scrollStepRow = 2;   // 垂直滚动2格
    const scrollStepCol = 1;   // 水平滚动1格
    const colStepCount = Math.floor(Math.abs(logicalDeltaX) / threshold);
    const rowStepCount = Math.floor(Math.abs(logicalDeltaY) / threshold);
    const frozenRows = sheet._frozenRows || 0;
    const frozenCols = sheet._frozenCols || 0;

    let needRender = false;

    // 水平滑动
    if (Math.abs(logicalDeltaX) > Math.abs(logicalDeltaY) && colStepCount > 0) {
        let newColStart;
        if (logicalDeltaX < 0) {
            // 向左滑动，视图向右移动（往后找列）
            newColStart = sheet._findNextShowCol(this.touchStartInfo.scrollX, scrollStepCol);
            if (newColStart !== sheet.vi.colArr[0]) {
                sheet.vi.colArr[0] = newColStart;
                needRender = true;
            }
        } else {
            // 向右滑动，视图向左移动（往前找列）
            newColStart = sheet._findBeforeShowCol(this.touchStartInfo.scrollX, scrollStepCol);
            newColStart = Math.max(newColStart, frozenCols); // 冻结限制
            if (newColStart !== sheet.vi.colArr[0]) {
                sheet.vi.colArr[0] = newColStart;
                needRender = true;
            }
        }
        // 更新起始位置
        if (needRender) {
            this.touchStartInfo.x = touch.clientX;
            this.touchStartInfo.scrollX = sheet.vi.colArr[0];
        }
    } else if (rowStepCount > 0) {
        // 垂直滑动
        let newRowStart;
        if (logicalDeltaY < 0) {
            // 向上滑动，视图向下移动（往后找行）
            newRowStart = sheet._findNextShowRow(this.touchStartInfo.scrollY, scrollStepRow);
            if (newRowStart !== sheet.vi.rowArr[0]) {
                sheet.vi.rowArr[0] = newRowStart;
                needRender = true;
            }
        } else {
            // 向下滑动，视图向上移动（往前找行）
            newRowStart = sheet._findBeforeShowRow(this.touchStartInfo.scrollY, scrollStepRow);
            newRowStart = Math.max(newRowStart, frozenRows); // 冻结限制
            if (newRowStart !== sheet.vi.rowArr[0]) {
                sheet.vi.rowArr[0] = newRowStart;
                needRender = true;
            }
        }
        // 更新起始位置
        if (needRender) {
            this.touchStartInfo.y = touch.clientY;
            this.touchStartInfo.scrollY = sheet.vi.rowArr[0];
        }
    }

    // 3. 真实渲染时重置CSS偏移
    if (needRender) {
        if (this.canvasContainer) {
            this.canvasContainer.style.transform = ''; // 清除CSS变换
        }
        this.r();
        this.updateScrollBar();
        this.touchStartInfo.lastRenderTime = now; // 更新渲染时间
    }
}

// 触屏滑动 - 触摸结束
export function canvasTouchEnd(event) {
    if (!this.touchStartInfo) return;

    // 清除CSS变换并进行最终渲染
    if (this.canvasContainer) {
        this.canvasContainer.style.transform = '';
    }

    // 最终渲染确保位置准确
    if (this.touchStartInfo.moved) {
        this.r();
        this.updateScrollBar();
    }

    // 清理状态
    this.touchStartInfo = null;

    // 如果输入框处于编辑状态，失焦
    if (this.inputEditing) {
        const blurEvent = new Event('blur');
        this.input.dispatchEvent(blurEvent);
    }
}

function dHandleResize(event, info, direction) {
    const { left, top, sheet } = info;
    let lastTime = 0;

    // 记录原始的起始和结束单元格信息
    const originalArea = { start: { ...this.activeDrawing.area.s }, end: { ...this.activeDrawing.area.e } };

    // 获取原始四个角的信息
    const startInfo = sheet.getCellInViewInfo(originalArea.start.r, originalArea.start.c);
    const endInfo = sheet.getCellInViewInfo(originalArea.end.r, originalArea.end.c);

    // 记录原始矩形的边界
    const originalBounds = {
        left: startInfo.x + originalArea.start.offsetX,
        top: startInfo.y + originalArea.start.offsetY,
        right: endInfo.x + originalArea.end.offsetX,
        bottom: endInfo.y + originalArea.end.offsetY
    };

    this.activeDrawing.active = true;
    const activeDrawing = this.activeDrawing
    downBeforeFuns.push(() => activeDrawing.active = false)

    const mouseMoveHandler = (moveEvent) => {
        // 节流处理
        const now = Date.now();
        if (now - lastTime < 16) return; // 约60fps
        lastTime = now;

        // 获取当前鼠标位置（相对于sheet的坐标）
        const moveX = Math.max(sheet.indexWidth, this.toLogical(moveEvent.clientX - left));
        const moveY = Math.max(sheet.headHeight, this.toLogical(moveEvent.clientY - top));

        // 根据拖拽方向计算新的边界
        const newBounds = { ...originalBounds };

        switch (direction) {
            case 'left-top':
                newBounds.left = moveX;
                newBounds.top = moveY;
                break;
            case 'right-top':
                newBounds.right = moveX;
                newBounds.top = moveY;
                break;
            case 'left-bottom':
                newBounds.left = moveX;
                newBounds.bottom = moveY;
                break;
            case 'right-bottom':
                newBounds.right = moveX;
                newBounds.bottom = moveY;
                break;
            case 'left':
                newBounds.left = moveX;
                break;
            case 'right':
                newBounds.right = moveX;
                break;
            case 'top':
                newBounds.top = moveY;
                break;
            case 'bottom':
                newBounds.bottom = moveY;
                break;
        }

        // 确保矩形有效（宽高为正）
        const rectBounds = {
            left: Math.min(newBounds.left, newBounds.right),
            top: Math.min(newBounds.top, newBounds.bottom),
            right: Math.max(newBounds.left, newBounds.right),
            bottom: Math.max(newBounds.top, newBounds.bottom)
        };

        // 确保最小尺寸
        const minSize = 20;
        if (rectBounds.right - rectBounds.left < minSize) {
            if (direction.includes('left')) {
                rectBounds.left = rectBounds.right - minSize;
            } else {
                rectBounds.right = rectBounds.left + minSize;
            }
        }
        if (rectBounds.bottom - rectBounds.top < minSize) {
            if (direction.includes('top')) {
                rectBounds.top = rectBounds.bottom - minSize;
            } else {
                rectBounds.bottom = rectBounds.top + minSize;
            }
        }

        // 获取新的起始单元格索引
        const startCellIndex = this.getCellIndexByPosition(rectBounds.left, rectBounds.top);

        // 获取起始单元格的视图信息
        const newStartCellInfo = sheet.getCellInViewInfo(startCellIndex.r, startCellIndex.c, false);

        // 计算起始单元格的偏移量
        const offsetX = rectBounds.left - newStartCellInfo.x;
        const offsetY = rectBounds.top - newStartCellInfo.y;

        // 更新绘制参数
        this.activeDrawing.startCell = { r: startCellIndex.r, c: startCellIndex.c };
        this.activeDrawing.offsetX = offsetX;
        this.activeDrawing.offsetY = offsetY;
        this.activeDrawing.width = rectBounds.right - rectBounds.left;
        this.activeDrawing.height = rectBounds.bottom - rectBounds.top;

        this.r('s');
    };

    // 添加事件监听器
    this.handleLayer.addEventListener('mousemove', mouseMoveHandler);

    window.addEventListener('mouseup', () => {
        this.handleLayer.removeEventListener('mousemove', mouseMoveHandler);
        this.disableMoveCheck = false;
        // 图表/形状需要重新渲染（尺寸改变影响布局），图片不需要（自动缩放）
        if (this.activeDrawing.type != 'image') {
            this.activeDrawing.drawCache = null;
        }
        this.r('s');
    }, { once: true });
}


function areaSel(start, left, top, width, height) {

    this.disableMoveCheck = true;
    let edgeCheckInterval = null; // 移动事件
    const sheet = this.activeSheet;
    const p = sheet.protection;
    const canSelectLocked = p.enabled ? p.selectLockedCells : true;
    const canSelectUnlocked = p.enabled ? p.selectUnlockedCells : true;
    let lastSelectToastAt = 0;
    const toastSelectOnce = (msg) => {
        const now = Date.now();
        if (now - lastSelectToastAt < 500) return;
        lastSelectToastAt = now;
        sheet.SN?.Utils?.toast?.(msg);
    };
    const getSelectionBlockReason = (area) => {
        if (!p.enabled) return '';
        if (!canSelectLocked && !canSelectUnlocked) return sheet.SN.t('core.sheet.sheetProtection.message.selectionBlocked');
        if (!area) return '';
        const range = typeof area === 'string' ? sheet.rangeStrToNum(area) : area;
        if (!range?.s || !range?.e) return '';
        const sr = Math.min(range.s.r, range.e.r);
        const er = Math.max(range.s.r, range.e.r);
        const sc = Math.min(range.s.c, range.e.c);
        const ec = Math.max(range.s.c, range.e.c);
        for (let r = sr; r <= er; r++) {
            for (let c = sc; c <= ec; c++) {
                const cell = sheet.getCell(r, c);
                const locked = p.isCellLocked(cell);
                if (locked && !canSelectLocked) return sheet.SN.t('core.sheet.sheetProtection.message.cannotSelectLocked');
                if (!locked && !canSelectUnlocked) return sheet.SN.t('core.sheet.sheetProtection.message.cannotSelectUnlocked');
            }
        }
        return '';
    };
    const blockSelection = (area) => {
        const reason = getSelectionBlockReason(area);
        if (!reason) return false;
        toastSelectOnce(reason);
        return true;
    };
    const cloneArea = (area) => area ? { s: { r: area.s.r, c: area.s.c }, e: { r: area.e.r, c: area.e.c } } : null;
    let lastAllowedArea = sheet.activeAreas.length ? cloneArea(sheet.activeAreas[sheet.activeAreas.length - 1]) : null;
    const onMouseMove = (moveEvent) => {

        const moveInfo = this.getCellIndex(moveEvent); // 当前鼠标所在的单元格索引
        const aLength = sheet.activeAreas.length;
        const lastArea = sheet.activeAreas[aLength - 1];

        if (aLength === 0) {
            const newArea = { s: { ...start }, e: { c: moveInfo.c, r: moveInfo.r } };
            const expandedArea = sheet._getExpandMaxArea(newArea);
            if (blockSelection(expandedArea)) return;
            sheet.activeAreas.push(expandedArea); // 添加最大选中区域
            lastAllowedArea = cloneArea(expandedArea);
            return this.r('s');
        } else if (start.c == -1) { // 选行
            if (lastArea.e.r == moveInfo.r) return
            moveInfo.c = sheet.colCount - 1
            const nextArea = adjustRange({ s: { r: start.r, c: 0 }, e: moveInfo });
            if (blockSelection(nextArea)) return;
            sheet.activeAreas[aLength - 1] = nextArea;
            lastAllowedArea = cloneArea(nextArea);
            return this.r('s');
        } else if (start.r == -1) { // 选列
            if (lastArea.e.c == moveInfo.c) return
            moveInfo.r = sheet.rowCount - 1
            const nextArea = adjustRange({ s: { r: 0, c: start.c }, e: moveInfo });
            if (blockSelection(nextArea)) return;
            sheet.activeAreas[aLength - 1] = nextArea;
            lastAllowedArea = cloneArea(nextArea);
            return this.r('s');
        } else { // 选数据
            const _moveInfo = moveInfo
            if (_moveInfo.r == -1) _moveInfo.r = sheet.vi.rowArr[0]
            if (_moveInfo.c == -1) _moveInfo.c = sheet.vi.colArr[0]
            // 标准先后顺序
            const temporaryArea = adjustRange({ s: { ...start }, e: moveInfo })
            const newArea = sheet._getExpandMaxArea(temporaryArea); // 更新最大选中区域
            if (blockSelection(newArea)) return;
            sheet.activeAreas[aLength - 1] = newArea;
            if (lastArea.e.c != newArea.e.c || lastArea.e.r != newArea.e.r || lastArea.s.c != newArea.s.c || lastArea.s.r != newArea.s.r) {
                this.r('s');
            }
            lastAllowedArea = cloneArea(newArea);
        }

        const zoom = this.activeSheet?.zoom ?? 1;
        const mouseX = moveEvent.clientX - left;
        const mouseY = moveEvent.clientY - top;
        const headerWidth = sheet.indexWidth * zoom;
        const headerHeight = sheet.headHeight * zoom;
        const viewWidth = sheet.vi.w * zoom;
        const viewHeight = sheet.vi.h * zoom;
        const isEdge = {
            l: mouseX <= 30 + headerWidth,
            r: mouseX >= width - 30 || mouseX > viewWidth,
            t: mouseY <= 30 + headerHeight,
            b: mouseY >= height - 30 || mouseY > viewHeight
        };
        clearInterval(edgeCheckInterval);

        edgeCheckInterval = setInterval(() => { // 通过改变vi起始才能改变视图
            const a = sheet.activeAreas;
            if (!a.length) return;
            const viewCol = sheet.vi.colArr
            const viewRow = sheet.vi.rowArr
            if (isEdge.l) {
                viewCol[0] = sheet._findBeforeShowCol(viewCol[0], 1);
                a[a.length - 1].s.c < a[a.length - 1].e.c ? a[a.length - 1].s.c = Math.max(viewCol[0], 0) : a[a.length - 1].e.c = Math.max(viewCol[0], 0)
            } else if (isEdge.r && sheet.vi.w * zoom > this.width) {
                viewCol[0] = viewCol[1] ?? sheet.vi.side.r;
                a[a.length - 1].e.c = Math.max(viewCol[viewCol.length - 1], 0)
            } else if (isEdge.t) {
                viewRow[0] = sheet._findBeforeShowRow(viewRow[0], 2);
                a[a.length - 1].s.r < a[a.length - 1].e.r ? a[a.length - 1].s.r = Math.max(viewRow[0], 0) : a[a.length - 1].e.r = Math.max(viewRow[0], 0)
            } else if (isEdge.b && sheet.vi.h * zoom > this.height) {
                viewRow[0] = viewRow[2] ?? viewRow[1] ?? sheet.vi.side.b;
                a[a.length - 1].s.r > a[a.length - 1].e.r ? a[a.length - 1].s.r = Math.max(viewRow[viewRow.length - 1], 0) : a[a.length - 1].e.r = Math.max(viewRow[viewRow.length - 1], 0);
            } else {
                return
            }
            const currentArea = a[a.length - 1];
            if (blockSelection(currentArea)) {
                if (lastAllowedArea) {
                    a[a.length - 1] = cloneArea(lastAllowedArea);
                }
                this.r();
                return;
            }
            lastAllowedArea = cloneArea(currentArea);
            this.r();
        }, 100);

    };

    this.handleLayer.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', () => {
        this.handleLayer.removeEventListener('mousemove', onMouseMove);
        this.disableMoveCheck = false;
        clearInterval(edgeCheckInterval); // 停止边缘检测
        if (!this.activeSheet._brush.area) return;

        // 格式刷操作 - 使用 Sheet.applyBrush 方法（支持平铺复制）
        const targetArea = this.activeSheet.activeAreas[this.activeSheet.activeAreas.length - 1];
        this.activeSheet.applyBrush(targetArea);
        this.r();
    }, { once: true });

}

function getMouseInRC(x, y, event, sheet) {
    if (x < sheet.indexWidth || y < sheet.headHeight) {
        const cellIndex = this.getCellIndex(event);
        const cellInfo = sheet.getCellInViewInfo(cellIndex.r, cellIndex.c);
        if (y < sheet.headHeight && (Math.abs(cellInfo.x - x) <= 5 || Math.abs(cellInfo.x + sheet.getCol(cellIndex.c).width - x) <= 5)) {
            this.moveAreaName = 'row-col-resize'
            return 'ew-resize';
        } else if (x < sheet.indexWidth && (Math.abs(cellInfo.y - y) <= 5 || Math.abs(cellInfo.y + sheet.getRow(cellIndex.r).height - y) <= 5)) {
            this.moveAreaName = 'row-col-resize'
            return 'ns-resize';
        } else {
            return null
        }
    }
}

function getMouseInDrawing(x, y) {
    const drawings = this.activeSheet.Drawing.getAll();
    const len = drawings.length;

    for (let i = len - 1; i >= 0; i--) {
        const d = drawings[i];
        if (!d.position || !d.inView) continue;

        const { x: dx, y: dy, w: dw, h: dh } = d.position;

        // 将鼠标坐标转换到图纸的本地坐标系（反向旋转）
        let lx = x, ly = y;
        if (d._rotation) {
            const cx = dx + dw / 2;
            const cy = dy + dh / 2;
            const rad = -d._rotation * Math.PI / 180;
            const cos = Math.cos(rad);
            const sin = Math.sin(rad);
            const ddx = x - cx;
            const ddy = y - cy;
            lx = cx + ddx * cos - ddy * sin;
            ly = cy + ddx * sin + ddy * cos;
        }

        // 如果图纸激活，先检测控制点（因为控制点在边界外也能点击）
        if (d.active) {
            // 预计算常用值
            const halfW = dw >> 1; // 位运算代替除法
            const halfH = dh >> 1;
            const x2 = dx + dw;
            const y2 = dy + dh;
            const midX = dx + halfW;
            const midY = dy + halfH;

            // 旋转手柄检测（图片 + 非连接线形状）
            if (d.type === 'image' || (d.type === 'shape' && !d.isConnector)) {
                const handleY = dy - 30;
                if (lx >= midX - 6 && lx <= midX + 6 && ly >= handleY - 6 && ly <= handleY + 6) {
                    this.activeDrawing = d;
                    this.moveAreaName = 'd-rotate';
                    return 'grab';
                }
            }

            // 控制点检测 - 8x8的正方形，中心在边界上
            // 左上角
            if (lx >= dx - 4 && lx <= dx + 4 && ly >= dy - 4 && ly <= dy + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-left-top';
                return 'nw-resize';
            }
            // 右上角
            if (lx >= x2 - 4 && lx <= x2 + 4 && ly >= dy - 4 && ly <= dy + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-right-top';
                return 'ne-resize';
            }
            // 左下角
            if (lx >= dx - 4 && lx <= dx + 4 && ly >= y2 - 4 && ly <= y2 + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-left-bottom';
                return 'sw-resize';
            }
            // 右下角
            if (lx >= x2 - 4 && lx <= x2 + 4 && ly >= y2 - 4 && ly <= y2 + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-right-bottom';
                return 'se-resize';
            }

            // 边中点检测
            // 上边中点
            if (lx >= midX - 4 && lx <= midX + 4 && ly >= dy - 4 && ly <= dy + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-top';
                return 'n-resize';
            }
            // 下边中点
            if (lx >= midX - 4 && lx <= midX + 4 && ly >= y2 - 4 && ly <= y2 + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-bottom';
                return 's-resize';
            }
            // 左边中点
            if (lx >= dx - 4 && lx <= dx + 4 && ly >= midY - 4 && ly <= midY + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-left';
                return 'w-resize';
            }
            // 右边中点
            if (lx >= x2 - 4 && lx <= x2 + 4 && ly >= midY - 4 && ly <= midY + 4) {
                this.activeDrawing = d;
                this.moveAreaName = 'd-right';
                return 'e-resize';
            }
        }

        // 快速边界检测 - 在图纸内部（使用本地坐标）
        if (lx >= dx && lx <= dx + dw && ly >= dy && ly <= dy + dh) {
            this.activeDrawing = d;
            this.moveAreaName = 'drawing';
            return 'move';
        }
    }

    return null;
}

function getDrawingByPoint(x, y) {
    const drawings = this.activeSheet.Drawing.getAll();
    for (let i = drawings.length - 1; i >= 0; i--) {
        const d = drawings[i];
        if (!d.position || !d.inView) continue;
        const { x: dx, y: dy, w: dw, h: dh } = d.position;

        let lx = x;
        let ly = y;
        if (d._rotation) {
            const cx = dx + dw / 2;
            const cy = dy + dh / 2;
            const rad = -d._rotation * Math.PI / 180;
            const cos = Math.cos(rad);
            const sin = Math.sin(rad);
            const ddx = x - cx;
            const ddy = y - cy;
            lx = cx + ddx * cos - ddy * sin;
            ly = cy + ddx * sin + ddy * cos;
        }

        if (lx >= dx && lx <= dx + dw && ly >= dy && ly <= dy + dh) {
            return d;
        }
    }
    return null;
}

function ensureTextBoxEditorState() {
    const existed = this._textBoxEditorState;
    if (existed?.editor?.isConnected) return existed;

    const host = this.SN.containerDom.querySelector('.sn-top-layer');
    if (!host) return null;

    const editor = document.createElement('div');
    editor.className = 'sn-textbox-editor';
    editor.contentEditable = 'true';
    editor.spellcheck = false;
    editor.style.display = 'none';
    editor.style.pointerEvents = 'none';
    host.appendChild(editor);

    const state = existed ?? {};
    state.editor = editor;
    state.active = false;
    state.drawing = null;
    state.originalText = '';
    state.closing = false;
    this._textBoxEditorState = state;

    this._syncTextBoxEditorPosition = () => syncTextBoxEditorPosition.call(this);
    this._finishTextBoxEditing = (commit = true) => finishTextBoxInlineEditing.call(this, commit);

    editor.addEventListener('mousedown', (event) => event.stopPropagation());
    editor.addEventListener('dblclick', (event) => event.stopPropagation());
    editor.addEventListener('keydown', (event) => {
        if (!state.active) return;
        event.stopPropagation();

        if (event.key === 'Escape') {
            event.preventDefault();
            finishTextBoxInlineEditing.call(this, false);
            this.input.focus();
            return;
        }

        if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
            event.preventDefault();
            finishTextBoxInlineEditing.call(this, true);
            this.input.focus();
            return;
        }

        if (event.key === 'Tab') {
            event.preventDefault();
            finishTextBoxInlineEditing.call(this, true);
            this.input.focus();
        }
    });
    editor.addEventListener('blur', () => {
        if (!state.active || state.closing) return;
        finishTextBoxInlineEditing.call(this, true);
    });

    return state;
}

function placeCaretToEnd(editor) {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(editor);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
}

function getEditorText(editor) {
    return String(editor?.innerText ?? '')
        .replace(/\r\n/g, '\n')
        .replace(/\u00a0/g, ' ')
        .replace(/\n$/, '');
}

function mapTextAlign(style = {}, isVertical = false) {
    if (isVertical) {
        const verticalAlignMap = { l: 'end', r: 'start', ctr: 'center' };
        return verticalAlignMap[style.textAlign] || 'start';
    }
    const horizontalAlignMap = { l: 'left', r: 'right', ctr: 'center' };
    return horizontalAlignMap[style.textAlign] || 'left';
}

function syncTextBoxEditorPosition() {
    const state = this._textBoxEditorState;
    const editor = state?.editor;
    if (!editor) return;

    const drawing = state.drawing;
    if (!state.active || !drawing?.position) {
        editor.style.display = 'none';
        editor.style.pointerEvents = 'none';
        return;
    }

    if (!drawing.inView) {
        editor.style.display = 'none';
        editor.style.pointerEvents = 'none';
        return;
    }

    const { x, y, w, h } = drawing.position;
    const style = drawing.shapeStyle || {};
    const textDirection = style.textDirection === 'vertical' ? 'vertical' : 'horizontal';
    const zoom = this.activeSheet?.zoom ?? 1;
    const fontSize = (style.fontSize || 11) * zoom;
    const strokeColor = style.stroke && style.stroke !== 'none' ? style.stroke : '#6493f0';
    const background = style.fill && style.fill !== 'none' ? style.fill : 'transparent';
    const textAlign = mapTextAlign(style, textDirection === 'vertical');

    editor.style.display = 'block';
    editor.style.pointerEvents = 'auto';
    editor.style.left = `${this.toPhysical(x)}px`;
    editor.style.top = `${this.toPhysical(y)}px`;
    editor.style.width = `${Math.max(12, this.toPhysical(w))}px`;
    editor.style.height = `${Math.max(12, this.toPhysical(h))}px`;
    editor.style.fontSize = `${fontSize}px`;
    editor.style.lineHeight = `${fontSize * 1.2}px`;
    editor.style.fontFamily = style.fontFamily || '宋体';
    editor.style.color = style.textColor || '#000000';
    editor.style.backgroundColor = background;
    editor.style.borderColor = strokeColor;
    editor.style.textAlign = textAlign;
    editor.style.writingMode = textDirection === 'vertical' ? 'vertical-rl' : 'horizontal-tb';
    editor.style.textOrientation = textDirection === 'vertical' ? 'upright' : 'mixed';
    editor.style.transformOrigin = 'center center';
    editor.style.transform = drawing._rotation ? `rotate(${drawing._rotation}deg)` : 'none';
}

function startTextBoxInlineEditing(drawing) {
    if (!drawing) return false;
    const state = ensureTextBoxEditorState.call(this);
    if (!state?.editor) return false;

    if (state.active && state.drawing && state.drawing !== drawing) {
        finishTextBoxInlineEditing.call(this, true);
    }

    state.active = true;
    state.drawing = drawing;
    state.originalText = drawing.shapeText || '';
    state.editor.innerText = state.originalText;

    this.activeDrawing = drawing;
    drawing.active = true;

    syncTextBoxEditorPosition.call(this);

    requestAnimationFrame(() => {
        if (!state.active || state.drawing !== drawing) return;
        state.editor.focus();
        placeCaretToEnd(state.editor);
    });

    return true;
}

function finishTextBoxInlineEditing(commit = true) {
    const state = this._textBoxEditorState;
    const editor = state?.editor;
    if (!state?.active || !editor || state.closing) return false;

    state.closing = true;
    const drawing = state.drawing;
    if (commit && drawing) {
        const nextText = getEditorText(editor);
        if (drawing.shapeText !== nextText) {
            drawing.shapeText = nextText;
            drawing.drawCache = null;
        }
    }

    state.active = false;
    state.drawing = null;
    state.originalText = '';
    editor.style.display = 'none';
    editor.style.pointerEvents = 'none';
    editor.style.transform = 'none';

    if (document.activeElement === editor) {
        editor.blur();
    }

    state.closing = false;
    this.r('s');
    return true;
}

// 区域索引排序
function adjustRange(range) {
    const { s, e } = range;

    const minC = Math.min(s.c, e.c);
    const minR = Math.min(s.r, e.r);
    const maxC = Math.max(s.c, e.c);
    const maxR = Math.max(s.r, e.r);

    return {
        s: { c: minC, r: minR },
        e: { c: maxC, r: maxR }
    };
}

// 将两个区域合并（**不考虑合并**）
function mergeExcelRanges(range1, range2) {
    // 确定新区域的起始行和列
    const startRow = Math.min(range1.s.r, range2.s.r, range1.e.r, range2.e.r);
    const startCol = Math.min(range1.s.c, range2.s.c, range1.e.c, range2.e.c);

    // 确定新区域的结束行和列
    const endRow = Math.max(range1.s.r, range2.s.r, range1.e.r, range2.e.r);
    const endCol = Math.max(range1.s.c, range2.s.c, range1.e.c, range2.e.c);

    // 返回合并后的区域索引
    return { s: { c: startCol, r: startRow }, e: { c: endCol, r: endRow } };
}

// 获取填充区域
function getPaddingArea(cellIndex) {
    const { s, e } = this.activeSheet.activeAreas[this.activeSheet.activeAreas.length - 1]; // 表格开始和结束的索引
    const { c, r } = cellIndex; // 鼠标所在的单元格索引
    // 检查是否在区域内
    if (c >= s.c && c <= e.c && r >= s.r && r <= e.r) return null;
    let paddingArea = { s: { c: c, r: r }, e: { c: c, r: r } }; // 初始化填充区域
    // 判断填充方向并计算填充区域
    if (c < s.c || c > e.c) { // 水平填充
        paddingArea.s.c = Math.min(e.c + 1, c);
        paddingArea.e.c = Math.max(s.c - 1, c);
        paddingArea.s.r = s.r;
        paddingArea.e.r = e.r;
    } else if (r < s.r || r > e.r) { // 垂直填充
        paddingArea.s.r = Math.min(e.r + 1, r);
        paddingArea.e.r = Math.max(s.r - 1, r);
        paddingArea.s.c = s.c;
        paddingArea.e.c = e.c;
    }
    return paddingArea;
}

// 获取鼠标在活动边框的信息，返回鼠标样式
function getMouseInBorder(x, y) {
    if (!this.activeSheet.protection.isSelectionVisible()) return null;
    const { x: rectX, y: rectY, w: rectW, h: rectH } = this.activeBorderInfo; // 矩形信息
    const borderWidth = 2; // 边框粗细
    const handleSize = 10; // 填充柄的大小
    // 判断鼠标是否在填充柄上
    if (x >= rectX + rectW + borderWidth / 2 - handleSize / 2 && x <= rectX + rectW + borderWidth / 2 + handleSize / 2 &&
        y >= rectY + rectH + borderWidth / 2 - handleSize / 2 && y <= rectY + rectH + borderWidth / 2 + handleSize / 2) {
        this.moveAreaName = 'fill-handle'
        return 'crosshair'
    }
    // 判断鼠标是否在边框上
    if (x >= rectX - borderWidth / 2 && x <= rectX + rectW + borderWidth / 2 &&
        y >= rectY - borderWidth / 2 && y <= rectY + rectH + borderWidth / 2 &&
        !(x > rectX + borderWidth / 2 && x < rectX + rectW - borderWidth / 2 &&
            y > rectY + borderWidth / 2 && y < rectY + rectH - borderWidth / 2)) {
        this.moveAreaName = 'active-border'
        return 'move'
    }
    return null;
}

// 移动新区域计算
function updateRegionBounds(region, oldCell, newCell, boundaries = this.activeSheet.vi.side) {

    region = region[region.length - 1];
    // 计算行和列的移动差异
    const rowDiff = newCell.r - oldCell.r;
    const colDiff = newCell.c - oldCell.c;
    // 创建一个新的区域对象来避免直接修改原始对象
    let newRegion = {
        s: { r: region.s.r + rowDiff, c: region.s.c + colDiff },
        e: { r: region.e.r + rowDiff, c: region.e.c + colDiff }
    };

    // 确保新的区域索引在边界内
    if (newRegion.s.r < boundaries.t) {
        const cha = boundaries.t - newRegion.s.r
        newRegion.s.r = boundaries.t
        newRegion.e.r += cha
    }
    if (newRegion.s.c < boundaries.l) {
        const cha = boundaries.l - newRegion.s.c
        newRegion.s.c = boundaries.l
        newRegion.e.c += cha
    }
    if (newRegion.e.r > boundaries.b) {
        const cha = newRegion.e.r - boundaries.b
        newRegion.e.r = boundaries.b
        newRegion.s.r -= cha
    }
    if (newRegion.e.c > boundaries.r) {
        const cha = newRegion.e.c - boundaries.r
        newRegion.e.c = boundaries.r
        newRegion.s.c -= cha
    }

    // 返回更新后的区域
    return newRegion;

}
