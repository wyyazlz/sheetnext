/**
 * Event manager
 * Initializes and manages event listeners.
 */

import { getColorSelectHTML } from './helpers.js';

/**
 * Initialize all event listeners.
 */
export function initEventListeners() {
    // Paste events
    const promptInput = this.SN.containerDom.querySelector('.sn-prompt-input');
    if (promptInput) {
        promptInput.addEventListener('paste', (event) => {
            const clipboardData = event.clipboardData || window.clipboardData;
            if (clipboardData) this.SN.AI._handlePastedImage(clipboardData);
        });
    }

    // Upload events
    const uploadInput = this.SN.containerDom.querySelector('.sn-upload');
    if (uploadInput) {
        uploadInput.addEventListener('change', async (event) => {
            this.SN.IO.import(event.target.files[0]);
        });
    }

    // Zoom control
    const zoomCtrl = this.SN.containerDom.querySelector('.sn-zoom-ctrl');
    if (zoomCtrl) {
        const SN = this.SN;
        const zoomOutBtn = zoomCtrl.querySelector('[data-icon="minus"]')?.closest('.sn-sheets-btn');
        const zoomInBtn = zoomCtrl.querySelector('[data-icon="plus"]')?.closest('.sn-sheets-btn');
        const zoomVal = zoomCtrl.querySelector('.sn-zoom-val');
        zoomOutBtn?.addEventListener('click', () => SN.activeSheet?.zoomOut());
        zoomInBtn?.addEventListener('click', () => SN.activeSheet?.zoomIn());
        zoomVal?.addEventListener('click', () => SN.Action.zoomCustom());
    }
}

/**
 * Initialize color components.
 * @param {Element} [scope] Optional scope.
 */
export function initColorComponents(scope) {
    const html = getColorSelectHTML(this.SN);
    const ns = this.SN.ns;
    const container = scope || this.SN.containerDom;
    container.querySelectorAll('.sn-color-select').forEach(item => {
        if (item.dataset.colorInit) return;
        item.innerHTML = html;
        item.dataset.colorInit = 'true';
        const colorInput = item.querySelector('.sn-color-input');
        if (colorInput) {
            const onclickAttr = item.getAttribute('onclick') || '';
            let changeHandler = null;
            if (onclickAttr.includes('fontColorBtn')) {
                changeHandler = (e) => { e.stopPropagation(); this.SN.Action.fontColorChange(e); };
            } else if (onclickAttr.includes('fillColorBtn')) {
                changeHandler = (e) => { e.stopPropagation(); this.SN.Action.fillColorChange(e); };
            } else if (onclickAttr.includes('borderColorBtn')) {
                changeHandler = (e) => { e.stopPropagation(); this.SN.Action.borderColorChange(e); };
            }
            if (changeHandler) {
                colorInput.addEventListener('input', changeHandler);
            }
            colorInput.addEventListener('click', (e) => e.stopPropagation());
        }
    });
}

/**
 * Remove loading animation.
 */
export function removeLoading() {
    const removeLoadingDiv = () => {
        const loadingDiv = this.SN.containerDom.querySelector('.sn-loading-div');
        if (loadingDiv) loadingDiv.remove();
    };

    if (document.readyState === 'complete') {
        removeLoadingDiv();
    } else {
        window.addEventListener('load', removeLoadingDiv);
    }
}
