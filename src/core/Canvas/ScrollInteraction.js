/**
 * Pixel-level wheel and touch scrolling for canvas.
 */

import { baseMousedown, canvasDblclick } from './InteractionHandler.js';

const INERTIA_MIN_VELOCITY = 0.018;
const INERTIA_MAX_VELOCITY = 3.2;
const PINCH_MIN_DISTANCE = 24;
const PINCH_ZOOM_EPSILON = 0.004;
const TAP_MAX_INTERVAL = 320;
const TAP_MAX_DISTANCE = 24;
const TOUCH_SELECTION_HANDLE_HIT_SIZE = 30;
const TOUCH_MOUSE_IGNORE_MS = 450;

function clampVelocity(value) {
    return Math.max(-INERTIA_MAX_VELOCITY, Math.min(INERTIA_MAX_VELOCITY, value));
}

function setTouchMouseIgnore(canvas) {
    canvas._ignoreTouchMouseUntil = performance.now() + TOUCH_MOUSE_IGNORE_MS;
}

function getTouchPoint(event, fallback) {
    return event.changedTouches?.[0] || event.touches?.[0] || fallback || null;
}

function createTouchMouseEvent(touch, sourceEvent, type = 'mousedown') {
    return {
        type,
        clientX: touch.clientX,
        clientY: touch.clientY,
        button: 0,
        buttons: type === 'mouseup' ? 0 : 1,
        ctrlKey: !!sourceEvent.ctrlKey,
        shiftKey: !!sourceEvent.shiftKey,
        altKey: !!sourceEvent.altKey,
        metaKey: !!sourceEvent.metaKey,
        _snPointerType: 'touch',
        _snSkipFocus: true,
        preventDefault: () => sourceEvent.preventDefault?.(),
        stopPropagation: () => sourceEvent.stopPropagation?.()
    };
}

function createSyntheticMouseEvent(type, touch, sourceEvent) {
    const event = new MouseEvent(type, {
        bubbles: true,
        cancelable: true,
        view: window,
        clientX: touch.clientX,
        clientY: touch.clientY,
        button: 0,
        buttons: type === 'mouseup' ? 0 : 1,
        ctrlKey: !!sourceEvent.ctrlKey,
        shiftKey: !!sourceEvent.shiftKey,
        altKey: !!sourceEvent.altKey,
        metaKey: !!sourceEvent.metaKey
    });
    try {
        event._snPointerType = 'touch';
        event._snSkipFocus = true;
    } catch (e) { }
    return event;
}

function dispatchTouchMouseEvent(canvas, type, touch, sourceEvent) {
    const event = createSyntheticMouseEvent(type, touch, sourceEvent);
    canvas.handleLayer.dispatchEvent(event);
    if (type === 'mouseup') {
        window.dispatchEvent(createSyntheticMouseEvent(type, touch, sourceEvent));
    }
}

function isTouchInSelectionHandle(canvas, touch) {
    const frame = canvas.activeBorderInfo;
    if (!frame?.inView) return false;
    const { x, y } = canvas.getEventPosition(touch);
    const paneClip = frame.paneClip;
    if (paneClip?.empty || (paneClip && (
        x < paneClip.x || x > paneClip.x + paneClip.w ||
        y < paneClip.y || y > paneClip.y + paneClip.h
    ))) return false;
    const zoom = canvas.activeSheet?.zoom ?? 1;
    const hitSize = Math.max(10, TOUCH_SELECTION_HANDLE_HIT_SIZE / zoom);
    const centerX = frame.x + frame.w + 1;
    const centerY = frame.y + frame.h + 1;
    return Math.abs(x - centerX) <= hitSize / 2 && Math.abs(y - centerY) <= hitSize / 2;
}

function cloneArea(area) {
    return area ? { s: { r: area.s.r, c: area.s.c }, e: { r: area.e.r, c: area.e.c } } : null;
}

function normalizeArea(area) {
    return {
        s: {
            r: Math.min(area.s.r, area.e.r),
            c: Math.min(area.s.c, area.e.c)
        },
        e: {
            r: Math.max(area.s.r, area.e.r),
            c: Math.max(area.s.c, area.e.c)
        }
    };
}

function isSameArea(a, b) {
    return a?.s?.r === b?.s?.r && a?.s?.c === b?.s?.c && a?.e?.r === b?.e?.r && a?.e?.c === b?.e?.c;
}

function clampCellToSheet(sheet, cell) {
    const fallbackRow = sheet.vi.rowArr?.[0] ?? 0;
    const fallbackCol = sheet.vi.colArr?.[0] ?? 0;
    const row = Number.isFinite(cell?.r) && cell.r >= 0 ? cell.r : fallbackRow;
    const col = Number.isFinite(cell?.c) && cell.c >= 0 ? cell.c : fallbackCol;
    return {
        r: Math.max(0, Math.min(sheet.rowCount - 1, row)),
        c: Math.max(0, Math.min(sheet.colCount - 1, col))
    };
}

function getSelectionBlockReason(sheet, area) {
    const p = sheet.protection;
    if (!p.enabled) return '';
    if (!p.selectLockedCells && !p.selectUnlockedCells) return sheet.SN.t('core.sheet.sheetProtection.message.selectionBlocked');
    const sr = Math.min(area.s.r, area.e.r);
    const er = Math.max(area.s.r, area.e.r);
    const sc = Math.min(area.s.c, area.e.c);
    const ec = Math.max(area.s.c, area.e.c);
    for (let r = sr; r <= er; r++) {
        for (let c = sc; c <= ec; c++) {
            const cell = sheet.getCell(r, c);
            const locked = p.isCellLocked(cell);
            if (locked && !p.selectLockedCells) return sheet.SN.t('core.sheet.sheetProtection.message.cannotSelectLocked');
            if (!locked && !p.selectUnlockedCells) return sheet.SN.t('core.sheet.sheetProtection.message.cannotSelectUnlocked');
        }
    }
    return '';
}

function canSelectArea(canvas, sheet, area) {
    const reason = getSelectionBlockReason(sheet, area);
    if (!reason) return true;
    const now = Date.now();
    if (!canvas._touchSelectionToastAt || now - canvas._touchSelectionToastAt > 500) {
        canvas._touchSelectionToastAt = now;
        sheet.SN?.Utils?.toast?.(reason);
    }
    return false;
}

function startTouchSelection(canvas, touch, event) {
    const sheet = canvas.activeSheet;
    const areas = sheet.activeAreas;
    const lastIndex = Math.max(0, areas.length - 1);
    const activeCell = clampCellToSheet(sheet, sheet.activeCell || canvas.getCellIndex(touch));
    const sourceArea = areas[lastIndex] || { s: { ...activeCell }, e: { ...activeCell } };
    const normalized = normalizeArea(sourceArea);

    event.preventDefault();
    setTouchMouseIgnore(canvas);
    canvas.moveAreaName = null;
    canvas.mpBorder = {};
    canvas.touchStartInfo = null;
    canvas.disableMoveCheck = true;
    canvas._lastTouchTap = null;
    canvas._touchSelectionInfo = {
        anchor: { r: normalized.s.r, c: normalized.s.c },
        lastIndex,
        lastArea: cloneArea(sourceArea),
        lastX: touch.clientX,
        lastY: touch.clientY
    };
    if (!areas.length) sheet.activeAreas = [cloneArea(sourceArea)];
    if (canvas.inputEditing) canvas.updInputValue();
}

function updateTouchSelection(canvas, touch) {
    const info = canvas._touchSelectionInfo;
    const sheet = canvas.activeSheet;
    if (!info || !sheet) return;
    info.lastX = touch.clientX;
    info.lastY = touch.clientY;

    const target = clampCellToSheet(sheet, canvas.getCellIndex(touch));
    const nextArea = sheet._getExpandMaxArea(normalizeArea({ s: info.anchor, e: target }));
    if (!canSelectArea(canvas, sheet, nextArea)) return;

    const index = Math.min(info.lastIndex, Math.max(0, sheet.activeAreas.length - 1));
    if (isSameArea(sheet.activeAreas[index], nextArea)) return;
    sheet.activeAreas[index] = nextArea;
    info.lastArea = cloneArea(nextArea);
    canvas.r('s');
}

function handleTouchTap(canvas, event, info) {
    const touch = getTouchPoint(event, { clientX: info.lastX, clientY: info.lastY });
    if (!touch) return;

    event.preventDefault();
    setTouchMouseIgnore(canvas);

    const tapEvent = createTouchMouseEvent(touch, event);
    const cellIndex = canvas.getCellIndex(tapEvent);
    if (canvas.inputEditing) canvas.updInputValue();
    canvas.moveAreaName = null;
    baseMousedown.call(canvas, tapEvent);
    dispatchTouchMouseEvent(canvas, 'mouseup', touch, event);

    if (cellIndex.r < 0 || cellIndex.c < 0) {
        canvas._lastTouchTap = null;
        return;
    }

    const now = performance.now();
    const lastTap = canvas._lastTouchTap;
    const isSameCell = lastTap?.sheet === canvas.activeSheet &&
        lastTap.cell?.r === cellIndex.r &&
        lastTap.cell?.c === cellIndex.c;
    const isClose = lastTap &&
        Math.hypot(touch.clientX - lastTap.x, touch.clientY - lastTap.y) <= TAP_MAX_DISTANCE;

    if (isSameCell && isClose && now - lastTap.time <= TAP_MAX_INTERVAL) {
        canvas._lastTouchTap = null;
        canvasDblclick.call(canvas, createTouchMouseEvent(touch, event, 'dblclick'));
        return;
    }

    canvas._lastTouchTap = {
        time: now,
        x: touch.clientX,
        y: touch.clientY,
        cell: cellIndex,
        sheet: canvas.activeSheet
    };
}

function clampZoom(value) {
    return Math.min(4, Math.max(0.1, value));
}

function getPinchTouchInfo(canvas, event) {
    if (event.touches.length < 2) return null;
    const rect = canvas.handleLayer.getBoundingClientRect();
    const first = event.touches[0];
    const second = event.touches[1];
    const firstX = first.clientX - rect.left;
    const firstY = first.clientY - rect.top;
    const secondX = second.clientX - rect.left;
    const secondY = second.clientY - rect.top;
    return {
        centerX: (firstX + secondX) / 2,
        centerY: (firstY + secondY) / 2,
        distance: Math.hypot(second.clientX - first.clientX, second.clientY - first.clientY)
    };
}

function setSheetZoom(canvas, sheet, nextZoom) {
    const normalized = clampZoom(nextZoom);
    const oldZoom = sheet.zoom ?? 1;
    if (Math.abs(oldZoom - normalized) < PINCH_ZOOM_EPSILON) return false;
    if (sheet.SN.Event.hasListeners('beforeZoomChange')) {
        const beforeEvent = sheet.SN.Event.emit('beforeZoomChange', { sheet, oldZoom, newZoom: normalized });
        if (beforeEvent.canceled) return false;
    }
    sheet._zoom = normalized;
    canvas.applyZoom();
    sheet.SN._updateZoomDisplay();
    if (sheet.SN.Event.hasListeners('afterZoomChange')) {
        sheet.SN.Event.emit('afterZoomChange', { sheet, oldZoom, newZoom: normalized });
    }
    return true;
}

function startPinchZoom(canvas, event) {
    const sheet = canvas.activeSheet;
    if (!sheet) return false;
    if (!sheet.vi) sheet._updView();
    const info = getPinchTouchInfo(canvas, event);
    if (!info || info.distance < PINCH_MIN_DISTANCE) return false;

    canvas._stopScrollInertia?.();
    canvas.touchStartInfo = null;

    const zoom = sheet.zoom ?? 1;
    const centerLogicalX = info.centerX / zoom;
    const centerLogicalY = info.centerY / zoom;
    const bodyLeft = sheet.vi.fw ?? sheet.indexWidth;
    const bodyTop = sheet.vi.fh ?? sheet.headHeight;
    const scrollLeft = sheet.vi.scrollLeft ?? sheet.getScrollLeft(sheet.vi.colArr?.[0] ?? 0);
    const scrollTop = sheet.vi.scrollTop ?? sheet.getScrollTop(sheet.vi.rowArr?.[0] ?? 0);

    canvas.pinchStartInfo = {
        distance: info.distance,
        zoom,
        anchorX: scrollLeft + Math.max(0, centerLogicalX - bodyLeft),
        anchorY: scrollTop + Math.max(0, centerLogicalY - bodyTop)
    };
    return true;
}

function movePinchZoom(canvas, event) {
    if (!canvas.pinchStartInfo && !startPinchZoom(canvas, event)) return;
    const sheet = canvas.activeSheet;
    const start = canvas.pinchStartInfo;
    const info = getPinchTouchInfo(canvas, event);
    if (!sheet || !start || !info || info.distance < PINCH_MIN_DISTANCE) return;

    event.preventDefault();
    const nextZoom = clampZoom(start.zoom * (info.distance / start.distance));
    const zoomChanged = setSheetZoom(canvas, sheet, nextZoom);
    const activeZoom = sheet.zoom ?? nextZoom;
    const centerLogicalX = info.centerX / activeZoom;
    const centerLogicalY = info.centerY / activeZoom;
    const bodyLeft = sheet.vi?.fw ?? sheet.indexWidth;
    const bodyTop = sheet.vi?.fh ?? sheet.headHeight;
    const nextLeft = start.anchorX - Math.max(0, centerLogicalX - bodyLeft);
    const nextTop = start.anchorY - Math.max(0, centerLogicalY - bodyTop);
    canvas.scrollToPixels(nextLeft, nextTop, { fast: true, force: zoomChanged });
}

export function canvasWheel(event) {
    if (this.disableScroll) return;
    this._stopScrollInertia?.();

    if (event.ctrlKey || event.metaKey) {
        event.preventDefault();
        const current = this.activeSheet?.zoom ?? 1;
        const delta = event.deltaY < 0 ? 0.1 : -0.1;
        const next = Math.round((current + delta) * 100) / 100;
        if (this.activeSheet) this.activeSheet.zoom = next;
        return;
    }

    let deltaX = event.deltaX || 0;
    let deltaY = event.deltaY || 0;
    if (event.shiftKey && Math.abs(deltaX) < Math.abs(deltaY)) {
        deltaX = deltaY;
        deltaY = 0;
    }

    const modeScale = event.deltaMode === 1 ? 40 : event.deltaMode === 2 ? this.height : 1;
    deltaX *= modeScale;
    deltaY *= modeScale;
    if (Math.abs(deltaX) < 0.01 && Math.abs(deltaY) < 0.01) return;

    event.preventDefault();
    const fast = Math.abs(deltaX) > this.width * 0.5 || Math.abs(deltaY) > this.height * 0.5;
    const zoom = this.activeSheet?.zoom ?? 1;
    this.scrollByPixels(deltaX / zoom, deltaY / zoom, { fast });

    if (this.inputEditing) {
        // 公式录入中允许滚动查看/选取引用区域，不提交编辑（与 Excel 一致）
        if (this.formulaEditor?.active) {
            this.formulaEditor.onViewScrolled();
        } else {
            const blurEvent = new Event('blur');
            this.input.dispatchEvent(blurEvent);
        }
    }
}

export function canvasTouchStart(event) {
    if (this.disableScroll) return;
    if (this._touchSelectionInfo) {
        event.preventDefault();
        return;
    }

    if (event.touches.length === 2) {
        event.preventDefault();
        startPinchZoom(this, event);
        return;
    }

    if (event.touches.length !== 1) return;
    if (this.disableMoveCheck) return;
    this._stopScrollInertia?.();

    const touch = event.touches[0];
    const sheet = this.activeSheet;
    if (!sheet) return;
    if (!sheet.vi) sheet._updView();

    if (isTouchInSelectionHandle(this, touch)) {
        startTouchSelection(this, touch, event);
        return;
    }

    event.preventDefault();
    this.moveAreaName = null;
    const now = performance.now();
    this.touchStartInfo = {
        x: touch.clientX,
        y: touch.clientY,
        lastX: touch.clientX,
        lastY: touch.clientY,
        time: now,
        lastTime: now,
        scrollLeft: sheet.vi.scrollLeft ?? sheet.getScrollLeft(sheet.vi.colArr?.[0] ?? 0),
        scrollTop: sheet.vi.scrollTop ?? sheet.getScrollTop(sheet.vi.rowArr?.[0] ?? 0),
        moved: false,
        velocityX: 0,
        velocityY: 0
    };
}

export function canvasTouchMove(event) {
    if (this.disableScroll) return;
    if (this._touchSelectionInfo) {
        if (event.touches.length !== 1) return;
        event.preventDefault();
        setTouchMouseIgnore(this);
        updateTouchSelection(this, event.touches[0]);
        return;
    }

    if (event.touches.length === 2) {
        event.preventDefault();
        movePinchZoom(this, event);
        return;
    }

    if (this.pinchStartInfo) return;
    if (!this.touchStartInfo || event.touches.length !== 1) return;

    const touch = event.touches[0];
    const info = this.touchStartInfo;
    const deltaX = touch.clientX - info.x;
    const deltaY = touch.clientY - info.y;
    if (!info.moved && (Math.abs(deltaX) > 5 || Math.abs(deltaY) > 5)) {
        info.moved = true;
    }
    if (!info.moved) return;

    event.preventDefault();
    const now = performance.now();
    const frameDt = Math.max(1, now - info.lastTime);
    const sampleX = (touch.clientX - info.lastX) / frameDt;
    const sampleY = (touch.clientY - info.lastY) / frameDt;
    const weight = frameDt > 48 ? 1 : 0.35;
    info.velocityX = clampVelocity(info.velocityX * (1 - weight) + sampleX * weight);
    info.velocityY = clampVelocity(info.velocityY * (1 - weight) + sampleY * weight);
    info.lastX = touch.clientX;
    info.lastY = touch.clientY;
    info.lastTime = now;

    const zoom = this.activeSheet?.zoom ?? 1;
    this.scrollToPixels(
        info.scrollLeft - deltaX / zoom,
        info.scrollTop - deltaY / zoom,
        { fast: Math.abs(info.velocityX) > 1.2 || Math.abs(info.velocityY) > 1.2 }
    );
}

export function canvasTouchEnd(event) {
    if (this._touchSelectionInfo) {
        event.preventDefault();
        const info = this._touchSelectionInfo;
        const touch = getTouchPoint(event, { clientX: info.lastX, clientY: info.lastY });
        if (touch && event?.type !== 'touchcancel') updateTouchSelection(this, touch);
        this._touchSelectionInfo = null;
        this.touchStartInfo = null;
        this._lastTouchTap = null;
        this.disableMoveCheck = false;
        setTouchMouseIgnore(this);
        return;
    }

    if (this.pinchStartInfo) {
        this.pinchStartInfo = null;
        this.touchStartInfo = null;
        this.r();
        if (this.inputEditing) {
            const blurEvent = new Event('blur');
            this.input.dispatchEvent(blurEvent);
        }
        return;
    }

    if (!this.touchStartInfo) return;
    const info = this.touchStartInfo;
    this.touchStartInfo = null;

    if (info.moved && event?.type !== 'touchcancel') {
        const zoom = this.activeSheet?.zoom ?? 1;
        const idleTime = performance.now() - info.lastTime;
        const idleFactor = idleTime > 120 ? 0 : Math.pow(0.92, idleTime / 16.67);
        const velocityX = -info.velocityX * idleFactor / zoom;
        const velocityY = -info.velocityY * idleFactor / zoom;
        if (Math.abs(velocityX) > INERTIA_MIN_VELOCITY || Math.abs(velocityY) > INERTIA_MIN_VELOCITY) {
            this._startScrollInertia?.(velocityX, velocityY);
        } else {
            this.r();
        }
    } else if (info.moved) {
        this.r();
    } else if (event?.type !== 'touchcancel') {
        handleTouchTap(this, event, info);
        return;
    }

    if (this.inputEditing) {
        const blurEvent = new Event('blur');
        this.input.dispatchEvent(blurEvent);
    }
}
