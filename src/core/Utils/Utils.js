/** @typedef {import('../Workbook/Workbook.js').default} SheetNext */

export default class Utils {
    static _modalDragInitialized = false;

    /** @param {SheetNext} SN */
    constructor(SN) {
        /** @type {SheetNext} */
        this._SN = SN;
        /** @type {HTMLTextAreaElement} */
        this._textArea = document.createElement('textarea');
        this._enableModalDrag();
    }

    /** @param {any} obj @returns {Array<any>} Normalize any value to array. */
    objToArr(obj) {
        if (Array.isArray(obj)) return obj;
        if (obj !== null && obj !== undefined) return [obj];
        return [];
    }

    /** @param {string} rgb @returns {string} Convert rgb() string to hex color. */
    rgbToHex(rgb) {
        const [r, g, b] = rgb.match(/\d+/g).map(Number);
        return `#${((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase()}`;
    }

    /** @param {number} emu @returns {number} Convert EMU to pixels. */
    emuToPx(emu) {
        return Math.round(emu * 96 / 914400);
    }

    /** @param {number} px @returns {number} Convert pixels to EMU. */
    pxToEmu(px) {
        return Math.round(px * 914400 / 96);
    }

    /** @param {string} encodedString @returns {string} Decode HTML entities to plain text. */
    decodeEntities(encodedString) {
        this._textArea.innerHTML = encodedString;
        return this._textArea.value;
    }

    /** @param {Object} obj @param {Array<string>} keyOrder @returns {void} Reorder object keys in-place. */
    sortObjectInPlace(obj, keyOrder) {
        const tempObj = {};
        for (const key of keyOrder) {
            if (key in obj) {
                tempObj[key] = obj[key];
            }
        }
        for (const key in obj) {
            if (!(key in tempObj)) {
                tempObj[key] = obj[key];
            }
        }
        for (const key in obj) delete obj[key];
        for (const key in tempObj) obj[key] = tempObj[key];
    }

    /** @param {Function} func @param {number} wait @returns {Function} Create a debounced function. */
    debounce(func, wait) {
        let timeout;
        return function (...args) {
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, args), wait);
        };
    }

    /** @param {number} num @returns {string} */
    numToChar(num) {
        let columnName = '';
        while (num >= 0) {
            const remainder = num % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            num = Math.floor(num / 26) - 1;
        }
        return columnName;
    }

    /** @param {string} name @returns {number} */
    charToNum(name) {
        let number = 0;
        for (let i = 0; i < name.length; i++) {
            number = number * 26 + (name.charCodeAt(i) - 64);
        }
        return number - 1;
    }

    /** @param {{s:{r:number,c:number},e:{r:number,c:number}}} obj @param {boolean} [absolute=false] @returns {string} */
    rangeNumToStr(obj, absolute = false) {
        const startCol = this.numToChar(obj.s.c);
        const startRow = obj.s.r + 1;
        const endCol = this.numToChar(obj.e.c);
        const endRow = obj.e.r + 1;
        if (absolute) {
            return `$${startCol}$${startRow}:$${endCol}$${endRow}`;
        }
        return `${startCol}${startRow}:${endCol}${endRow}`;
    }

    /** @param {string} cellStr @returns {{c:number,r:number}|undefined} */
    cellStrToNum(cellStr) {
        if (!cellStr) return;
        const match = cellStr.match(/([A-Z]+)(\d+)/i);
        if (!match) return;
        return { c: this.charToNum(match[1].toUpperCase()), r: parseInt(match[2], 10) - 1 };
    }

    /** @param {{c:number,r:number}} cellNum @returns {string} */
    cellNumToStr(cellNum) {
        return this.numToChar(cellNum.c) + (cellNum.r + 1);
    }

    /** @param {string} formula @param {number} originalRow @param {number} originalCol @param {number} targetRow @param {number} targetCol @returns {string} Build fill formula by shifting relative refs only. */
    getFillFormula(formula, originalRow, originalCol, targetRow, targetCol) {
        return formula.replace(/(\$?[A-Z]+)(\$?\d+)/g, (_, colPart, rowPart) => {
            let col = this.charToNum(colPart.replace('$', ''));
            let row = parseInt(rowPart.replace('$', ''), 10) - 1;
            if (!colPart.includes('$')) col = col - originalCol + targetCol;
            if (!rowPart.includes('$')) row = row - originalRow + targetRow;
            return `${colPart.includes('$') ? '$' : ''}${this.numToChar(col)}${rowPart.includes('$') ? '$' : ''}${row + 1}`;
        });
    }

    /** @param {any} message @returns {void} Show toast message. */
    toast(message) {
        let toastContainer = this._SN.containerDom.querySelector('.sn-toast-container');
        if (!toastContainer) {
            toastContainer = document.createElement('div');
            toastContainer.classList.add('sn-toast-container');
            this._SN.containerDom.appendChild(toastContainer);
        }
        const messageDiv = document.createElement('div');
        messageDiv.classList.add('toast-item');
        const rawMessage = String(message);
        messageDiv.textContent = this._SN.t(rawMessage);
        toastContainer.appendChild(messageDiv);

        requestAnimationFrame(() => {
            messageDiv.classList.add('toast-show');
        });

        setTimeout(() => {
            messageDiv.classList.add('toast-hide');
            setTimeout(() => {
                if (messageDiv.parentNode === toastContainer) {
                    toastContainer.removeChild(messageDiv);
                }
            }, 300);
        }, 3000);
    }

    /** @param {string} imageUrl @returns {Promise<string|undefined>} Fetch image URL and return base64 data URL. */
    async downloadImageToBase64(imageUrl) {
        try {
            const response = await fetch(imageUrl);
            if (!response.ok) throw new Error(`Failed to fetch image: ${response.status} ${response.statusText}`);
            const blob = await response.blob();
            return await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.toast(this._SN.t('core.utils.utils.toast.imageDownloadFailed', { message }));
            return undefined;
        }
    }

    /** @param {Object} [options={}] @returns {Promise<{bodyEl: HTMLElement}|null>} Open common modal panel. */
    modal(options = {}) {
        const defaultTitle = this._SN.t('common.notice');
        const closeText = this._SN.t('common.close');
        const {
            titleKey = '',
            title = defaultTitle,
            contentKey = '',
            content = '',
            bodyKey = '',
            body = '',
            maxWidth = '460px',
            maxHeight = '',
            confirmKey = '',
            confirmText = null,
            cancelKey = '',
            cancelText = null,
            bodyClass = '',
            onOpen = null,
            onConfirm = null
        } = options;
        const bodyContent = content || body;
        const bodyTranslationKey = contentKey || bodyKey;
        const toFallbackString = (value) => {
            if (value === null || value === undefined) return '';
            return String(value);
        };
        const resolveText = (key, fallback) => {
            const fallbackText = toFallbackString(fallback);
            return key ? this._SN.t(key) : fallbackText;
        };
        const translatedTitle = resolveText(titleKey, title);
        const bodyFallback = toFallbackString(bodyContent);
        const bodyTemplate = bodyTranslationKey ? this._SN.t(bodyTranslationKey) : bodyFallback;
        const translatedBody = bodyTemplate;
        const translatedConfirmText = confirmText === null && !confirmKey ? '' : resolveText(confirmKey, confirmText);
        const translatedCancelText = cancelText === null && !cancelKey ? '' : resolveText(cancelKey, cancelText);

        return new Promise((resolve) => {
            let modalOverlay = this._SN.containerDom.querySelector('.sn-modal-overlay');
            if (!modalOverlay) {
                modalOverlay = document.createElement('div');
                modalOverlay.className = 'sn-modal-overlay';
                modalOverlay.innerHTML = `
                    <div class="sn-modal">
                        <div class="sn-modal-header">
                            <h2 class="sn-modal-title"></h2>
                            <button class="sn-modal-close" aria-label="${closeText}"></button>
                        </div>
                        <div class="sn-modal-body"></div>
                        <div class="sn-modal-footer">
                            <button class="sn-modal-btn sn-modal-btn-cancel"></button>
                            <button class="sn-modal-btn sn-modal-btn-confirm"></button>
                        </div>
                    </div>
                `;
                this._SN.containerDom.appendChild(modalOverlay);
            }

            const modalEl = modalOverlay.querySelector('.sn-modal');
            const titleEl = modalOverlay.querySelector('.sn-modal-title');
            const bodyEl = modalOverlay.querySelector('.sn-modal-body');
            const closeBtn = modalOverlay.querySelector('.sn-modal-close');
            const cancelBtn = modalOverlay.querySelector('.sn-modal-btn-cancel');
            const confirmBtn = modalOverlay.querySelector('.sn-modal-btn-confirm');

            const newCloseBtn = closeBtn.cloneNode(true);
            const newCancelBtn = cancelBtn.cloneNode(true);
            const newConfirmBtn = confirmBtn.cloneNode(true);
            closeBtn.replaceWith(newCloseBtn);
            cancelBtn.replaceWith(newCancelBtn);
            confirmBtn.replaceWith(newConfirmBtn);
            newCloseBtn.setAttribute('aria-label', closeText);

            titleEl.textContent = translatedTitle;
            bodyEl.innerHTML = translatedBody;
            bodyEl.className = `sn-modal-body${bodyClass ? ` ${bodyClass}` : ''}`;
            modalEl.style.maxWidth = maxWidth || '';
            modalEl.style.maxHeight = maxHeight || '';

            const footerEl = modalOverlay.querySelector('.sn-modal-footer');
            newConfirmBtn.textContent = translatedConfirmText || '';
            newCancelBtn.textContent = translatedCancelText || '';
            newConfirmBtn.style.display = translatedConfirmText ? '' : 'none';
            newCancelBtn.style.display = translatedCancelText ? '' : 'none';
            footerEl.style.display = translatedConfirmText || translatedCancelText ? '' : 'none';

            const cleanup = () => {
                modalOverlay.classList.remove('active', 'closing');
                titleEl.textContent = '';
                bodyEl.innerHTML = '';
                bodyEl.className = 'sn-modal-body';
                modalEl.style.maxWidth = '';
                modalEl.style.maxHeight = '';
                modalEl.style.transform = '';
                footerEl.style.display = '';
            };

            const closeModal = (isConfirm = false) => {
                modalOverlay.classList.add('closing');
                setTimeout(cleanup, 200);
                resolve(isConfirm ? { bodyEl } : null);
            };

            newCloseBtn.onclick = () => closeModal(false);
            newCancelBtn.onclick = () => closeModal(false);
            newConfirmBtn.onclick = async () => {
                if (onConfirm) {
                    try {
                        const result = await onConfirm(bodyEl);
                        if (result === false) return;
                    } catch (error) {
                        if (error) this.toast(error);
                        return;
                    }
                }
                closeModal(true);
            };

            requestAnimationFrame(() => {
                modalOverlay.classList.add('active');
                if (onOpen) onOpen(bodyEl);
            });
        });
    }

    _enableModalDrag() {
        if (Utils._modalDragInitialized) return;
        Utils._modalDragInitialized = true;

        let isDragging = false;
        let currentModal = null;
        let startX;
        let startY;
        let initialX;
        let initialY;
        let originalTransition = '';

        document.addEventListener('mousedown', (e) => {
            const header = e.target.closest('.sn-modal-header');
            if (!header || e.target.closest('.sn-modal-close')) return;

            const modal = header.closest('.sn-modal');
            if (!modal) return;

            isDragging = true;
            currentModal = modal;
            startX = e.clientX;
            startY = e.clientY;

            originalTransition = modal.style.transition;
            modal.style.transition = 'none';

            const match = modal.style.transform.match(/translate\(([^,]+)px,\s*([^)]+)px\)/);
            if (match) {
                initialX = parseFloat(match[1]);
                initialY = parseFloat(match[2]);
            } else {
                initialX = 0;
                initialY = 0;
            }

            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging || !currentModal) return;
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            currentModal.style.transform = `translate(${initialX + dx}px, ${initialY + dy}px)`;
        });

        document.addEventListener('mouseup', () => {
            if (currentModal && originalTransition !== undefined) {
                currentModal.style.transition = originalTransition;
            }
            isDragging = false;
            currentModal = null;
        });
    }
}
