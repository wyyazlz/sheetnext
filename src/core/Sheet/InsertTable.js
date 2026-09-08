/**
 * Insert template values, formatting and merges into a worksheet.
 */

const INSERT_TABLE_DEPRECATION_DATE = '2026-05-27';
const BATCH_CELL_THRESHOLD = 100000;
const SLICE_DURATION_MS = 12;
const PROGRESS_INTERVAL_MS = 100;
const CHECK_INTERVAL = 64;
const _pendingInsertions = new WeakMap();
let insertTableDeprecatedWarned = false;

function _yieldToBrowser() {
    if (globalThis.scheduler?.yield) return globalThis.scheduler.yield();
    return new Promise(resolve => setTimeout(resolve, 0));
}

/**
 * Insert a template, yielding with progress when its target exceeds 100,000 cells.
 * Await completion before reading the inserted range or performing dependent edits.
 * @param {(ICellConfig|string|number)[][]} arr - Template data array.
 * @param {CellRef} pos - Top-left insertion position.
 * @param {{align?:string,border?:boolean,width?:number,height?:number}} [options={}] - Template formatting.
 * @returns {Promise<RangeNum|undefined>} Inserted range, or undefined for empty/invalid input.
 */
export async function insertTemplate(arr, pos, options = {}) {
    if (!Array.isArray(arr) || !arr.length || !pos) return;
    if (typeof pos === 'string') pos = this.Utils.cellStrToNum(pos);
    if (!pos || !Number.isInteger(pos.r) || !Number.isInteger(pos.c) || pos.r < 0 || pos.c < 0) return;
    pos = { r: pos.r, c: pos.c };

    const SN = this.SN;
    // Serialize insertions per workbook because history and the modal are shared by its sheets.
    const previous = _pendingInsertions.get(SN);
    let release;
    const pending = new Promise(resolve => { release = resolve; });
    _pendingInsertions.set(SN, pending);

    let progress = null;
    let undoStarted = false;
    let sliceStart = performance.now();
    let lastProgress = 0;
    let lastStage = '';
    const checkpoint = (cellCount, percent, stage, params = {}) => {
        if (cellCount <= BATCH_CELL_THRESHOLD) return null;
        const first = !progress;
        if (first) progress = this.Utils._createProgressModal('core.sheet.insertTemplate.title');
        const now = performance.now();
        if (first || stage !== lastStage || now - lastProgress >= PROGRESS_INTERVAL_MS) {
            progress.update(percent, SN.t(`core.sheet.insertTemplate.${stage}`, params));
            lastProgress = now;
            lastStage = stage;
        }
        if (!first && now - sliceStart < SLICE_DURATION_MS) return null;
        return _yieldToBrowser().then(() => { sliceStart = performance.now(); });
    };

    try {
        if (previous) await previous;
        sliceStart = performance.now();
        const alignMap = { l: 'left', c: 'center', r: 'right' };
        const rowCount = arr.length;
        let maxCol = 0;
        let scanned = 0;

        // Grow the target-size lower bound while scanning, so preparation can also yield.
        for (let rIndex = 0; rIndex < rowCount; rIndex++) {
            const row = arr[rIndex] ?? [];
            let colIndex = 0;
            let skip = 0;
            for (let i = 0; i < row.length; i++) {
                if (skip > 0) {
                    skip--;
                } else {
                    const cell = row[i];
                    const mergeCols = cell?.cs ? cell.cs - 1 : (cell?.mr ?? 0);
                    colIndex += mergeCols + 1;
                    skip = mergeCols;
                }
                if (++scanned % CHECK_INTERVAL === 0) {
                    const pause = checkpoint(rowCount * Math.max(maxCol, colIndex), 5 * rIndex / rowCount, 'preparing');
                    if (pause) await pause;
                }
            }
            maxCol = Math.max(maxCol, colIndex);
            if ((!progress && rowCount * maxCol > BATCH_CELL_THRESHOLD) || rIndex % CHECK_INTERVAL === 0) {
                const pause = checkpoint(rowCount * maxCol, 5 * (rIndex + 1) / rowCount, 'preparing');
                if (pause) await pause;
            }
        }
        if (!maxCol) return;

        const totalCells = rowCount * maxCol;
        if (progress) {
            SN.UndoRedo.begin('insertTemplate');
            undoStarted = true;
        }
        const readyMerge = [];
        const borderStyle = { style: 'thin', color: '#000000' };
        let written = 0;
        for (let rIndex = 0; rIndex < rowCount; rIndex++) {
            const item = arr[rIndex] ?? [];
            for (let cIndex = 0; cIndex < maxCol; cIndex++) {
                const r = pos.r + rIndex;
                const c = pos.c + cIndex;
                const cell = this.getCell(r, c);

                if (options.border) cell.border = { top: borderStyle, left: borderStyle, bottom: borderStyle, right: borderStyle };
                if (options.align) cell.alignment = { vertical: 'center', horizontal: options.align };

                const tCell = item[cIndex] ?? {};
                if (Object.prototype.toString.call(tCell) !== '[object Object]') {
                    cell.editVal = tCell;
                } else {
                    cell.editVal = tCell.v;
                    if (tCell.fg) cell.fill = { fgColor: tCell.fg };
                    if (tCell.c) cell.font = { color: tCell.c };
                    if (tCell.a) cell.alignment = { horizontal: alignMap[tCell.a] };
                    if (tCell.b) cell.font = { bold: true };
                    if (tCell.s) cell.font = { size: Number(tCell.s) };

                    const mergeR = tCell.cs ? tCell.cs - 1 : (tCell.mr ?? 0);
                    const mergeB = tCell.rs ? tCell.rs - 1 : (tCell.mb ?? 0);
                    if (mergeR || mergeB) {
                        readyMerge.push({ s: { r, c }, e: { r: r + mergeB, c: c + mergeR } });
                    }
                }

                written++;
                if (progress && written % CHECK_INTERVAL === 0) {
                    const pause = checkpoint(totalCells, 5 + 90 * written / totalCells, 'writing', {
                        current: Math.floor(written / maxCol), total: rowCount
                    });
                    if (pause) await pause;
                }
            }
            if (item[0]?.h || options.height) this.getRow(pos.r + rIndex).height = item[0]?.h ?? options.height;
        }

        for (let index = 0; index < maxCol; index++) {
            const col = this.getCol(pos.c + index);
            const item = arr[0]?.[index] ?? {};
            const width = item.w ?? options.width;
            if (width) col.width = width;
            if (progress && index % CHECK_INTERVAL === 0) {
                const pause = checkpoint(totalCells, 95 + 2 * (index + 1) / maxCol, 'formatting');
                if (pause) await pause;
            }
        }

        for (let index = 0; index < readyMerge.length; index++) {
            this.mergeCells(readyMerge[index]);
            if (progress) {
                const pause = checkpoint(totalCells, 97 + 2 * (index + 1) / readyMerge.length, 'formatting');
                if (pause) await pause;
            }
        }

        if (progress) progress.update(100, SN.t('core.sheet.insertTemplate.done'));
        return { s: pos, e: { r: pos.r + rowCount - 1, c: pos.c + maxCol - 1 } };
    } finally {
        try {
            if (undoStarted) {
                SN.UndoRedo.commit();
                SN._r();
            }
        } finally {
            try {
                progress?.close();
            } finally {
                release();
                if (_pendingInsertions.get(SN) === pending) _pendingInsertions.delete(SN);
            }
        }
    }
}

/**
 * Insert a template from the given position.
 * @deprecated Use insertTemplate(arr, pos, options) instead. insertTable has ambiguous naming and will be deprecated after 2026-05-27.
 * @param {Array} arr - Template data array.
 * @param {Object|String} pos - Insert location.
 * @param {Object} options - Configure options.
 * @returns {Promise<RangeNum|undefined>} Inserted range after completion.
 */
export function insertTable(arr, pos, options = {}) {
    if (!insertTableDeprecatedWarned && typeof console !== 'undefined') {
        insertTableDeprecatedWarned = true;
        console.warn(`[SheetNext] sheet.insertTable() is kept for compatibility, but the name is ambiguous and will be deprecated after ${INSERT_TABLE_DEPRECATION_DATE}. Use sheet.insertTemplate() instead.`);
    }
    return insertTemplate.call(this, arr, pos, options);
}
