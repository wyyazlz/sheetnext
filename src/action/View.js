// 视图操作模块（缩放、冻结）

export function zoomCustom() {
    const SN = this.SN;
    const current = Math.round((SN.activeSheet?.zoom ?? 1) * 100);
    const content = `
        <div class="sn-zoom-dialog sn-numfmt-dialog">
            <div class="sn-zoom-row sn-numfmt-row">
                <label>Zoom Scale</label>
                <input type="number" class="sn-zoom-input" min="10" max="400" value="${current}">
                <span class="sn-zoom-suffix">%</span>
            </div>
        </div>
    `;
    SN.Utils.modal({
        titleKey: 'action.view.modal.m001.title', title: 'Custom Zoom',
        maxWidth:'300px',
        content,
        confirmKey: 'action.view.modal.m001.confirm', confirmText: 'Apply',
        cancelKey: 'action.view.modal.m001.cancel', cancelText: 'Cancel',
        onOpen: (bodyEl) => {
            const input = bodyEl.querySelector('.sn-zoom-input');
            if (!input) return;
            input.focus();
            input.select();
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    bodyEl.closest('.sn-modal')?.querySelector('.sn-modal-btn-confirm')?.click();
                }
            });
        },
        onConfirm: (bodyEl) => {
            const input = bodyEl.querySelector('.sn-zoom-input');
            const val = Number(input?.value);
            if (!Number.isFinite(val)) {
                SN.Utils.toast(SN.t('action.view.toast.pleaseEnterANumberBetween'));
                return false;
            }
            const percent = Math.min(400, Math.max(10, Math.round(val)));
            if (SN.activeSheet) SN.activeSheet.zoom = percent / 100;
            return true;
        }
    });
}

// ============================================================
// 冻结窗格操作
// ============================================================

// 冻结首行（当前视图的第一行，只冻结1行）
export function freezeRows(count) {
    const sheet = this.SN.activeSheet;
    const currentFirstRow = sheet.vi?.rowArr?.[0] ?? 0;
    // 冻结起始 = 当前视图首行
    sheet._freezeStartRow = currentFirstRow;
    // 冻结结束 = 起始 + count
    sheet.frozenRows = currentFirstRow + count;
    this.SN._r();
}

// 冻结首列（当前视图的第一列，只冻结1列）
export function freezeCols(count) {
    const sheet = this.SN.activeSheet;
    const currentFirstCol = sheet.vi?.colArr?.[0] ?? 0;
    sheet._freezeStartCol = currentFirstCol;
    sheet.frozenCols = currentFirstCol + count;
    this.SN._r();
}

// 同时冻结首行首列
export function freezeRowsAndCols(rows, cols) {
    const sheet = this.SN.activeSheet;
    const currentFirstRow = sheet.vi?.rowArr?.[0] ?? 0;
    const currentFirstCol = sheet.vi?.colArr?.[0] ?? 0;
    sheet._freezeStartRow = currentFirstRow;
    sheet._freezeStartCol = currentFirstCol;
    sheet.frozenRows = currentFirstRow + rows;
    sheet.frozenCols = currentFirstCol + cols;
    this.SN._r();
}

// 根据当前活动单元格冻结（冻结窗格）
// 冻结活动单元格上方和左侧的所有可见行列
export function freezeAtActiveCell() {
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];
    const row = area.s.r;
    const col = area.s.c;

    if (row === 0 && col === 0) {
        this.SN.Utils.toast(this.SN.t('action.view.toast.msg002'));
        return;
    }

    const currentFirstRow = sheet.vi?.rowArr?.[0] ?? 0;
    const currentFirstCol = sheet.vi?.colArr?.[0] ?? 0;

    // 冻结从当前视图首行/列开始，到活动单元格位置
    sheet._freezeStartRow = currentFirstRow;
    sheet._freezeStartCol = currentFirstCol;
    sheet.frozenRows = row;
    sheet.frozenCols = col;
    this.SN._r();
}

// 冻结至当前行（冻结活动单元格上方的可见行，当前行在非冻结区域）
export function freezeToCurrentRow() {
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];
    const row = area.s.r;

    if (row === 0) {
        this.SN.Utils.toast(this.SN.t('action.view.toast.msg003'));
        return;
    }

    const currentFirstRow = sheet.vi?.rowArr?.[0] ?? 0;
    sheet._freezeStartRow = currentFirstRow;
    sheet.frozenRows = row; // 冻结到当前行之前，当前行是非冻结区域的第一行
    this.SN._r();
}

// 冻结至当前列（冻结活动单元格左侧的可见列，当前列在非冻结区域）
export function freezeToCurrentCol() {
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];
    const col = area.s.c;

    if (col === 0) {
        this.SN.Utils.toast(this.SN.t('action.view.toast.msg004'));
        return;
    }

    const currentFirstCol = sheet.vi?.colArr?.[0] ?? 0;
    sheet._freezeStartCol = currentFirstCol;
    sheet.frozenCols = col; // 冻结到当前列之前，当前列是非冻结区域的第一列
    this.SN._r();
}

// 取消冻结
export function unfreeze() {
    const sheet = this.SN.activeSheet;
    sheet._freezeStartRow = 0;
    sheet._freezeStartCol = 0;
    sheet.frozenCols = 0;
    sheet.frozenRows = 0;
    this.SN._r();
}
