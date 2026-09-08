import getSvg from '../../assets/mainSvgs.js';
import { getRangePaneClip } from './helpers.js';

export default class HyperlinkPopover {
    constructor(canvas) {
        this._canvas = canvas;
        this._element = null;
        this._target = null;
        this._targetText = null;
        this._editButton = null;
        this._removeButton = null;
        this._sheet = null;
        this._cell = null;
        this._dismissedCell = null;
        this._suspended = false;
    }

    _getSelection() {
        const canvas = this._canvas;
        const sheet = canvas.activeSheet;
        if (!sheet?.activeCell || canvas.inputEditing || canvas.activeDrawing?.active ||
            !sheet.protection.isSelectionVisible()) return null;
        const active = sheet.getCell(sheet.activeCell.r, sheet.activeCell.c);
        const position = active.master || sheet.activeCell;
        const cell = sheet.getCell(position.r, position.c);
        const target = cell.hyperlink?.location || cell.hyperlink?.target;
        if (!target) return null;
        const area = cell.isMerged
            ? sheet.merges.find(merge => merge.s.r === position.r && merge.s.c === position.c)
            : { s: position, e: position };
        const selected = sheet.activeAreas;
        if (!area || selected.length !== 1) return null;
        const selection = selected[0];
        if (Math.min(selection.s.r, selection.e.r) !== area.s.r ||
            Math.max(selection.s.r, selection.e.r) !== area.e.r ||
            Math.min(selection.s.c, selection.e.c) !== area.s.c ||
            Math.max(selection.s.c, selection.e.c) !== area.e.c) return null;
        return { sheet, cell, area, target: String(target) };
    }

    _canEdit(sheet, cell) {
        return !this._canvas.SN.readOnly && (!sheet.protection.enabled ||
            (!sheet.protection.isCellLocked(cell) && sheet.protection.insertHyperlinks));
    }

    _create() {
        const element = document.createElement('div');
        element.className = 'sn-hyperlink-popover';
        element.setAttribute('role', 'group');
        element.hidden = true;
        const target = document.createElement('button');
        target.type = 'button';
        target.className = 'sn-hyperlink-target';
        target.setAttribute('role', 'link');
        target.innerHTML = getSvg('chaolianjie');
        const targetText = document.createElement('span');
        target.appendChild(targetText);
        const editButton = document.createElement('button');
        editButton.type = 'button';
        editButton.className = 'sn-hyperlink-action';
        editButton.innerHTML = getSvg('hyperlinkEdit');
        const removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.className = 'sn-hyperlink-action sn-hyperlink-remove';
        removeButton.innerHTML = getSvg('hyperlinkRemove');
        element.append(target, editButton, removeButton);
        element.querySelectorAll('svg').forEach(icon => icon.setAttribute('aria-hidden', 'true'));
        element.addEventListener('pointerdown', event => event.stopPropagation());
        element.addEventListener('mousedown', event => event.stopPropagation());
        element.addEventListener('dblclick', event => event.stopPropagation());
        element.addEventListener('keydown', event => {
            event.stopPropagation();
            if (event.key === 'Escape' || event.key === 'F6') {
                event.preventDefault();
                this._dismiss();
                this._canvas.input.focus({ preventScroll: true });
            }
        });
        target.addEventListener('click', () => this._activate('open'));
        editButton.addEventListener('click', () => this._activate('edit'));
        removeButton.addEventListener('click', () => this._activate('remove'));
        this._canvas.SN.containerDom.querySelector('.sn-top-layer').appendChild(element);
        this._element = element;
        this._target = target;
        this._targetText = targetText;
        this._editButton = editButton;
        this._removeButton = removeButton;
    }

    _hide() {
        if (this._element) this._element.hidden = true;
    }

    _dismiss() {
        this._dismissedCell = this._getSelection()?.cell || null;
        this._hide();
    }

    _reset() {
        this._dismissedCell = null;
        this._hide();
    }

    _handlePointerDown(event) {
        if (this._element?.contains(event.target)) return;
        if (event.target === this._canvas.handleLayer && event.button !== 2) this._reset();
        else this._dismiss();
    }

    _handleFocus(event) {
        if (event.target === this._canvas.input || this._element?.contains(event.target)) this._sync();
        else this._dismiss();
    }

    _focus() {
        this._dismissedCell = null;
        this._sync();
        if (!this._element || this._element.hidden) return false;
        this._target.focus({ preventScroll: true });
        return true;
    }

    _sync() {
        const canvas = this._canvas;
        const selection = this._getSelection();
        if (selection?.cell !== this._dismissedCell) this._dismissedCell = null;
        if (!selection || this._suspended || selection.cell === this._dismissedCell ||
            (document.activeElement !== canvas.input && !this._element?.contains(document.activeElement))) {
            this._hide();
            return;
        }
        const { sheet, cell, area, target } = selection;
        const info = sheet.getAreaInviewInfo(area);
        const clip = getRangePaneClip(canvas, sheet, area);
        if (!info.inView || clip.empty) return this._hide();
        const cellLeft = Math.max(info.x, clip.x);
        const cellTop = Math.max(info.y, clip.y);
        const cellRight = Math.min(info.x + info.w, clip.x + clip.w);
        const cellBottom = Math.min(info.y + info.h, clip.y + clip.h);
        if (cellRight <= cellLeft || cellBottom <= cellTop) return this._hide();
        const margin = 6;
        const leftEdge = canvas.toPhysical(sheet.indexWidth) + margin;
        const topEdge = canvas.toPhysical(sheet.headHeight) + margin;
        const rightEdge = canvas.toPhysical(canvas.viewWidth) - margin;
        const bottomEdge = canvas.toPhysical(canvas.viewHeight) - margin;
        if (rightEdge - leftEdge < 100 || bottomEdge - topEdge < 40) return this._hide();
        if (!this._element) this._create();
        const translate = key => canvas.SN.t(`core.canvas.hyperlinkPopover.${key}`);
        const canEdit = this._canEdit(sheet, cell);
        this._element.setAttribute('aria-label', translate('label'));
        this._targetText.textContent = target;
        this._target.title = cell.hyperlink.tooltip ? `${cell.hyperlink.tooltip}\n${target}` : target;
        this._target.setAttribute('aria-label', `${translate('open')}: ${target}`);
        this._editButton.title = translate('edit');
        this._editButton.setAttribute('aria-label', translate('edit'));
        this._editButton.hidden = !canEdit;
        this._removeButton.title = translate('remove');
        this._removeButton.setAttribute('aria-label', translate('remove'));
        this._removeButton.hidden = !canEdit;
        this._element.style.maxWidth = `${Math.min(360, rightEdge - leftEdge)}px`;
        this._element.hidden = false;
        const width = this._element.offsetWidth;
        const height = this._element.offsetHeight;
        const left = Math.max(leftEdge, Math.min(canvas.toPhysical(cellLeft), rightEdge - width));
        let top = canvas.toPhysical(cellBottom) + margin;
        if (top + height > bottomEdge) top = canvas.toPhysical(cellTop) - height - margin;
        top = Math.max(topEdge, Math.min(top, bottomEdge - height));
        this._element.style.left = `${left}px`;
        this._element.style.top = `${top}px`;
        this._sheet = sheet;
        this._cell = cell;
    }

    async _activate(action) {
        const selection = this._getSelection();
        if (!selection || selection.sheet !== this._sheet || selection.cell !== this._cell) {
            this._dismiss();
            return;
        }
        const { sheet, cell, area } = selection;
        if (action !== 'open' && !this._canEdit(sheet, cell)) return;
        this._dismiss();
        if (action === 'edit') {
            this._suspended = true;
            try {
                await this._canvas.SN.Action.insertHyperlink();
            } finally {
                this._suspended = false;
                this._canvas.input.focus({ preventScroll: true });
                this._canvas.r();
            }
        } else {
            if (action === 'open') sheet.hyperlinkJump();
            else sheet.clearRange(area, 'hyperlinks');
            this._canvas.input.focus({ preventScroll: true });
            this._canvas.r();
        }
    }
}
