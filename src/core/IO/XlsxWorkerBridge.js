import JSZip from "jszip";
import XlsxWorker from "./workers/xlsxWorker.js?worker&inline";

const PROGRESS_HIDE_DELAY = 260;
const IMPORT_READ_START = 4;
const IMPORT_READ_END = 12;

function _createWorker() {
    if (typeof Worker === 'undefined') return null;
    return new XlsxWorker();
}

function _runWorker(type, payload, transfer = [], onProgress = null) {
    return new Promise((resolve, reject) => {
        const worker = _createWorker();
        if (!worker) {
            const error = new Error('xlsx_worker_unavailable');
            error.workerUnavailable = true;
            reject(error);
            return;
        }

        worker.onmessage = (event) => {
            const message = event.data || {};
            if (message.type === 'progress') {
                onProgress?.(message.payload || {});
                return;
            }
            worker.terminate();
            if (message.type === 'success') {
                resolve(message.payload);
            } else {
                reject(new Error(message.error || 'xlsx_worker_failed'));
            }
        };

        worker.onerror = (event) => {
            worker.terminate();
            reject(new Error(event.message || 'xlsx_worker_error'));
        };

        try {
            worker.postMessage({ type, payload }, transfer);
        } catch (error) {
            worker.terminate();
            reject(error);
        }
    });
}

function _nextFrame() {
    return new Promise(resolve => {
        if (typeof requestAnimationFrame === 'function') {
            requestAnimationFrame(() => resolve());
        } else {
            setTimeout(resolve, 0);
        }
    });
}

function _formatPercent(value) {
    const pct = Math.max(0, Math.min(100, Math.round(Number(value) || 0)));
    return `${pct}%`;
}

function _progressText(SN, key, params = {}) {
    const path = `core.iO.iOProgress.${key}`;
    const text = SN.t(path, params);
    return text === path ? key : text;
}

function _scaleProgress(value, start, end) {
    const safeValue = Math.max(0, Math.min(100, Number(value) || 0));
    return start + ((end - start) * safeValue / 100);
}

function _createProgress(SN, titleKey) {
    const container = SN.containerDom;
    const host = container.querySelector('.sn-main') || container;
    let overlay = host.querySelector('.sn-xlsx-progress-overlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.className = 'sn-xlsx-progress-overlay';
        overlay.innerHTML = `
            <div class="sn-xlsx-progress-panel" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
                <div class="sn-xlsx-progress-header">
                    <div class="sn-xlsx-progress-title"></div>
                    <div class="sn-xlsx-progress-percent"></div>
                </div>
                <div class="sn-xlsx-progress-track">
                    <div class="sn-xlsx-progress-bar"></div>
                </div>
                <div class="sn-xlsx-progress-status"></div>
            </div>
        `;
        host.appendChild(overlay);
    }

    const titleEl = overlay.querySelector('.sn-xlsx-progress-title');
    const panelEl = overlay.querySelector('.sn-xlsx-progress-panel');
    const percentEl = overlay.querySelector('.sn-xlsx-progress-percent');
    const barEl = overlay.querySelector('.sn-xlsx-progress-bar');
    const statusEl = overlay.querySelector('.sn-xlsx-progress-status');
    let hideTimer = null;

    const update = (progress, statusKey, params = {}) => {
        if (hideTimer) {
            clearTimeout(hideTimer);
            hideTimer = null;
        }
        const safeProgress = Math.max(0, Math.min(100, Number(progress) || 0));
        titleEl.textContent = _progressText(SN, titleKey);
        percentEl.textContent = _formatPercent(safeProgress);
        panelEl.setAttribute('aria-valuenow', String(Math.round(safeProgress)));
        barEl.style.width = `${safeProgress}%`;
        statusEl.textContent = _progressText(SN, String(statusKey || '').split('.').pop(), params);
        overlay.classList.add('active');
    };

    const done = () => {
        update(100, 'done');
        hideTimer = setTimeout(() => {
            overlay.classList.remove('active');
        }, PROGRESS_HIDE_DELAY);
    };

    const fail = () => {
        overlay.classList.remove('active');
    };

    return { update, done, fail };
}

async function _readFileArrayBuffer(file, onProgress = null) {
    if (typeof FileReader === 'undefined') {
        const buffer = await file.arrayBuffer();
        onProgress?.({ progress: 100, loaded: file.size, total: file.size });
        return buffer;
    }

    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => resolve(event.target.result);
        reader.onerror = () => reject(reader.error || new Error('file_read_failed'));
        reader.onprogress = (event) => {
            if (!event.lengthComputable) return;
            onProgress?.({
                progress: Math.min(100, (event.loaded / Math.max(1, event.total)) * 100),
                loaded: event.loaded,
                total: event.total
            });
        };
        reader.readAsArrayBuffer(file);
    });
}

function _arrayBufferToBinaryString(buffer) {
    const view = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const chunkSize = 0x8000;
    const chunks = [];
    for (let i = 0; i < view.length; i += chunkSize) {
        chunks.push(String.fromCharCode.apply(null, view.subarray(i, i + chunkSize)));
    }
    return chunks.join('');
}

function _createTextEntry(name, text) {
    return {
        name,
        _data: text,
        asText() {
            return text;
        },
        asBinary() {
            return text;
        }
    };
}

function _createBinaryEntry(name, buffer) {
    let binaryCache = null;
    const view = new Uint8Array(buffer);
    return {
        name,
        _data: view,
        asText() {
            return _arrayBufferToBinaryString(view);
        },
        asBinary() {
            if (binaryCache === null) binaryCache = _arrayBufferToBinaryString(view);
            return binaryCache;
        },
        asUint8Array() {
            return view;
        },
        asArrayBuffer() {
            return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength);
        }
    };
}

function _loadXmlObjFromEntries(SN, entries, sheetRows = null, sheetRowMeta = null) {
    const xmlObj = Object.create(null);
    entries.forEach(entry => {
        if (!entry?.name) return;
        if (entry.type === 'text') {
            xmlObj[entry.name] = _createTextEntry(entry.name, entry.data || '');
        } else if (entry.data instanceof ArrayBuffer) {
            xmlObj[entry.name] = _createBinaryEntry(entry.name, entry.data);
        }
    });

    SN.Xml._styleCache = [];
    SN.Xml._sheetDataRowsByPath = sheetRows || Object.create(null);
    SN.Xml._sheetDataRowMetaByPath = sheetRowMeta || Object.create(null);
    SN.Xml.obj = xmlObj;
    SN._loadXmlObj();
}

function _importXlsxOnMain(SN, buffer) {
    const zip = new JSZip();
    zip.load(buffer);
    SN.Xml._styleCache = [];
    SN.Xml._sheetDataRowsByPath = Object.create(null);
    SN.Xml._sheetDataRowMetaByPath = Object.create(null);
    SN.Xml.obj = zip.files;
    SN._loadXmlObj();
}

function _normalizeTransferBuffer(data) {
    if (data instanceof ArrayBuffer) return data;
    if (ArrayBuffer.isView(data)) {
        if (data.byteOffset === 0 && data.byteLength === data.buffer.byteLength) return data.buffer;
        return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
    }
    return null;
}

export async function importXlsxWithWorker(io, file) {
    const SN = io._SN;
    const progress = _createProgress(SN, 'importTitle');
    progress.update(IMPORT_READ_START, 'readingFile');
    await _nextFrame();

    try {
        const buffer = await _readFileArrayBuffer(file, item => {
            progress.update(_scaleProgress(item.progress, IMPORT_READ_START, IMPORT_READ_END), 'readingFile', item);
        });

        progress.update(IMPORT_READ_END, 'openingPackage');
        await _nextFrame();

        try {
            const result = await _runWorker('import-xlsx', { buffer }, [buffer], item => {
                progress.update(item.progress ?? IMPORT_READ_END, item.stage || 'parsing', item);
            });
            progress.update(86, 'applyingData');
            await _nextFrame();
            _loadXmlObjFromEntries(SN, result?.entries || [], result?.sheetRows, result?.sheetRowMeta);
        } catch (error) {
            console.warn('[SheetNext] XLSX worker import failed, fallback to main thread import.', error);
            progress.update(16, 'openingPackage');
            await _nextFrame();
            const fallbackBuffer = error.workerUnavailable ? buffer : await _readFileArrayBuffer(file);
            _importXlsxOnMain(SN, fallbackBuffer);
        }

        progress.update(96, 'rendering');
        SN.workbookName = file?.name?.replace(/\.[^/.]+$/, '') || SN.t('workbook.defaultName');
        progress.done();
    } catch (error) {
        progress.fail();
        throw error;
    }
}

export async function exportXlsxWithWorker(io) {
    const SN = io._SN;
    const progress = _createProgress(SN, 'exportTitle');
    progress.update(4, 'collectingData');
    await _nextFrame();

    try {
        await io._exportXlsx({
            useWorkerZip: true,
            onProgress(item = {}) {
                progress.update(item.progress ?? 4, item.stage || 'collectingData', item);
            }
        });
        progress.done();
    } catch (error) {
        progress.fail();
        throw error;
    }
}

export async function zipXlsxEntriesWithWorker(entries, onProgress = null) {
    const transfer = [];
    const payloadEntries = entries.map(entry => {
        if (entry.type === 'binary') {
            const buffer = _normalizeTransferBuffer(entry.data);
            const transferBuffer = buffer ? buffer.slice(0) : null;
            if (transferBuffer) transfer.push(transferBuffer);
            return { name: entry.name, type: 'binary', data: transferBuffer };
        }
        return { name: entry.name, type: 'text', data: entry.data ?? '' };
    });

    const result = await _runWorker('zip-entries', { entries: payloadEntries }, transfer, onProgress);
    if (!(result?.buffer instanceof ArrayBuffer)) throw new Error('xlsx_worker_empty_export_result');
    return result.buffer;
}
