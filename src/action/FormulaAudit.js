/**
 * 公式审核模块
 * 提供公式追踪、错误检查、求值等功能
 */

// 箭头数据存储在 sheet 上
function getArrowData(sheet) {
    if (!sheet._traceArrows) {
        sheet._traceArrows = { precedents: [], dependents: [] };
    }
    return sheet._traceArrows;
}

/**
 * 解析公式中的单元格引用（仅当前工作表）
 * @param {string} formula - 公式字符串
 * @returns {Array} 引用的单元格列表 [{r, c}]
 */
function parseFormulaReferences(formula) {
    if (!formula || !formula.startsWith('=')) return [];

    const refs = [];
    // 匹配单元格引用：A1, $A$1, A1:B10 等（排除跨表引用）
    const cellPattern = /(?<![!'"])\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/gi;
    const matches = formula.match(cellPattern) || [];

    matches.forEach(match => {
        // 跳过跨表引用
        if (match.includes('!')) return;

        const clean = match.replace(/\$/g, '');
        const parts = clean.split(':');

        parts.forEach(part => {
            const colMatch = part.match(/^([A-Z]+)(\d+)$/i);
            if (colMatch) {
                const c = colMatch[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
                const r = parseInt(colMatch[2], 10) - 1;
                if (!refs.some(ref => ref.r === r && ref.c === c)) {
                    refs.push({ r, c });
                }
            }
        });
    });

    return refs;
}

/**
 * 追踪引用单元格（当前单元格引用了哪些单元格）
 */
export function tracePrecedents() {
    const sheet = this.SN.activeSheet;
    const { r, c } = sheet.activeCell;
    const cell = sheet.getCell(r, c);

    if (!cell?.editVal?.startsWith('=')) {
        this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.msg001'));
        return;
    }

    const refs = parseFormulaReferences(cell.editVal);
    if (refs.length === 0) {
        this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.msg002'));
        return;
    }

    const arrows = getArrowData(sheet);
    // 添加箭头：从被引用单元格指向当前单元格
    refs.forEach(ref => {
        const exists = arrows.precedents.some(
            a => a.from.r === ref.r && a.from.c === ref.c && a.to.r === r && a.to.c === c
        );
        if (!exists) {
            arrows.precedents.push({ from: ref, to: { r, c } });
        }
    });

    this.SN._r();
}

/**
 * 追踪从属单元格（哪些单元格引用了当前单元格）
 */
export function traceDependents() {
    const sheet = this.SN.activeSheet;
    const { r: activeR, c: activeC } = sheet.activeCell;
    const activeRef = this.SN.Utils.numToChar(activeC) + (activeR + 1);

    const arrows = getArrowData(sheet);
    const dependents = [];

    // 遍历所有单元格查找引用当前单元格的公式
    sheet.rows.forEach((row, ri) => {
        row.cells.forEach((cell, ci) => {
            if (!cell?.editVal?.startsWith('=')) return;
            if (ri === activeR && ci === activeC) return;

            const refs = parseFormulaReferences(cell.editVal);
            if (refs.some(ref => ref.r === activeR && ref.c === activeC)) {
                dependents.push({ r: ri, c: ci });
            }
        });
    });

    if (dependents.length === 0) {
        this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.msg003'));
        return;
    }

    // 添加箭头：从当前单元格指向依赖它的单元格
    dependents.forEach(dep => {
        const exists = arrows.dependents.some(
            a => a.from.r === activeR && a.from.c === activeC && a.to.r === dep.r && a.to.c === dep.c
        );
        if (!exists) {
            arrows.dependents.push({ from: { r: activeR, c: activeC }, to: dep });
        }
    });

    this.SN._r();
}

/**
 * 移除追踪箭头
 * @param {string} type - 'precedents' | 'dependents' | 'all'
 */
export function removeArrows(type = 'all') {
    const sheet = this.SN.activeSheet;
    const arrows = getArrowData(sheet);

    if (type === 'precedents' || type === 'all') {
        arrows.precedents = [];
    }
    if (type === 'dependents' || type === 'all') {
        arrows.dependents = [];
    }

    this.SN._r();
}

/**
 * 切换显示公式模式
 */
export function showFormulas() {
    const sheet = this.SN.activeSheet;
    sheet.showFormulas = !sheet.showFormulas;
    this.SN.Utils.toast(sheet.showFormulas
        ? this.SN.t('action.formulaAudit.toast.showFormulas')
        : this.SN.t('action.formulaAudit.toast.showCalculatedValues'));
    this.SN._r();
}

/**
 * 错误检查
 */
export function errorCheck() {
    const t = (key, params) => this.SN.t(key, params);
    const sheet = this.SN.activeSheet;
    const errors = [];
    const errorTypes = ['#REF!', '#NAME?', '#VALUE!', '#DIV/0!', '#NULL!', '#NUM!', '#N/A', '#CALC!'];

    sheet.rows.forEach((row, ri) => {
        row.cells.forEach((cell, ci) => {
            if (!cell) return;
            const val = cell.calcVal;
            if (typeof val === 'string' && errorTypes.includes(val)) {
                errors.push({
                    ref: this.SN.Utils.numToChar(ci) + (ri + 1),
                    r: ri,
                    c: ci,
                    error: val,
                    formula: cell.editVal || ''
                });
            }
        });
    });

    if (errors.length === 0) {
        this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.msg004'));
        return;
    }

    const listHtml = errors.map((e, i) => `
        <div class="sn-error-item${i === 0 ? ' active' : ''}" data-r="${e.r}" data-c="${e.c}">
            <span class="sn-error-ref">${e.ref}</span>
            <span class="sn-error-type">${e.error}</span>
            <span class="sn-error-formula" title="${e.formula}">${e.formula}</span>
        </div>
    `).join('');

    this.SN.Utils.modal({
        titleKey: 'action.formulaAudit.modal.m001.title',
        title: t('action.formulaAudit.modal.m001.title'),
        maxWidth: '400px',
        content: `
            <div class="sn-error-check">
                <div class="sn-error-summary">${t('action.formulaAudit.content.errorSummary', { count: errors.length })}</div>
                <div class="sn-error-list">${listHtml}</div>
            </div>
        `,
        confirmKey: 'action.formulaAudit.modal.m001.confirm',
        confirmText: t('action.formulaAudit.modal.m001.confirm'),
        onOpen: (bodyEl) => {
            const list = bodyEl.querySelector('.sn-error-list');
            list.addEventListener('click', (e) => {
                const item = e.target.closest('.sn-error-item');
                if (!item) return;
                list.querySelectorAll('.sn-error-item').forEach(el => el.classList.remove('active'));
                item.classList.add('active');
                // 跳转到错误单元格
                const r = parseInt(item.dataset.r, 10);
                const c = parseInt(item.dataset.c, 10);
                sheet.activeCell = { r, c };
                sheet.activeAreas = [{ s: { r, c }, e: { r, c } }];
                this.SN._r();
            });
        },
        onConfirm: () => true
    });
}

/**
 * 公式求值（简单展示）
 */
export function evaluateFormula() {
    const t = (key, params) => this.SN.t(key, params);
    const sheet = this.SN.activeSheet;
    const { r, c } = sheet.activeCell;
    const cell = sheet.getCell(r, c);

    if (!cell?.editVal?.startsWith('=')) {
        this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.msg001'));
        return;
    }

    const formula = cell.editVal;
    const refs = parseFormulaReferences(formula);

    // 构建求值步骤
    let steps = [];
    let evalFormula = formula;

    refs.forEach(ref => {
        const refCell = sheet.getCell(ref.r, ref.c);
        const refStr = this.SN.Utils.numToChar(ref.c) + (ref.r + 1);
        const refVal = refCell?.calcVal ?? 0;
        steps.push({ ref: refStr, value: refVal });
        // 替换公式中的引用为值（简单替换，可能不完全准确）
        const pattern = new RegExp(`\\$?${refStr.replace(/\$/g, '\\$?')}`, 'gi');
        evalFormula = evalFormula.replace(pattern, refVal);
    });

    const stepsHtml = steps.length > 0 ? steps.map(s => `
        <div class="sn-eval-step">
            <span class="sn-eval-ref">${s.ref}</span>
            <span class="sn-eval-arrow">→</span>
            <span class="sn-eval-val">${s.value}</span>
        </div>
    `).join('') : `<div class="sn-eval-empty">${t('action.formulaAudit.content.evaluate.noReferences')}</div>`;

    this.SN.Utils.modal({
        titleKey: 'action.formulaAudit.modal.m002.title',
        title: t('action.formulaAudit.modal.m002.title'),
        maxWidth: '360px',
        content: `
            <div class="sn-formula-eval">
                <div class="sn-eval-row">
                    <label>${t('action.formulaAudit.content.evaluate.formulaLabel')}</label>
                    <span class="sn-eval-formula">${formula}</span>
                </div>
                <div class="sn-eval-row">
                    <label>${t('action.formulaAudit.content.evaluate.referenceReplaceLabel')}</label>
                </div>
                <div class="sn-eval-steps">${stepsHtml}</div>
                <div class="sn-eval-row">
                    <label>${t('action.formulaAudit.content.evaluate.replacedLabel')}</label>
                    <span class="sn-eval-replaced">${evalFormula}</span>
                </div>
                <div class="sn-eval-row sn-eval-result-row">
                    <label>${t('action.formulaAudit.content.evaluate.resultLabel')}</label>
                    <span class="sn-eval-result">${cell.calcVal}</span>
                </div>
            </div>
        `,
        confirmKey: 'action.formulaAudit.modal.m002.confirm',
        confirmText: t('action.formulaAudit.modal.m002.confirm'),
        onConfirm: () => true
    });
}

/**
 * 重新计算工作簿
 */
export function calculateNow() {
    const startTime = performance.now();
    this.SN.sheets.forEach(sheet => {
        sheet.rows.forEach(row => {
            row.cells.forEach(cell => {
                if (cell?.editVal?.startsWith('=')) {
                    cell._calcVal = null; // 清除缓存强制重算
                }
            });
        });
    });
    this.SN._r();
    const elapsed = (performance.now() - startTime).toFixed(1);
    this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.recalculatedWorkbook', { elapsed }));
}

/**
 * 重新计算当前工作表
 */
export function calculateSheet() {
    const startTime = performance.now();
    const sheet = this.SN.activeSheet;
    sheet.rows.forEach(row => {
        row.cells.forEach(cell => {
            if (cell?.editVal?.startsWith('=')) {
                cell._calcVal = null;
            }
        });
    });
    this.SN._r();
    const elapsed = (performance.now() - startTime).toFixed(1);
    this.SN.Utils.toast(this.SN.t('action.formulaAudit.toast.recalculatedSheet', { elapsed }));
}

/**
 * 设置计算模式
 * @param {string} mode - 'auto' | 'manual'
 */
export function setCalcMode(mode) {
    this.SN.calcMode = mode;
    this.SN.Utils.toast(mode === 'manual' ? this.SN.t('action.formulaAudit.toast.switchToManual') : this.SN.t('action.formulaAudit.toast.switchToAuto'));
}

/**
 * 渲染追踪箭头（在 Canvas 渲染后调用）
 * @param {CanvasRenderingContext2D} ctx - 画布上下文
 * @param {Sheet} sheet - 工作表对象
 */
export function renderTraceArrows(ctx, sheet) {
    const arrows = sheet._traceArrows;
    if (!arrows) return;

    const dpr = window.devicePixelRatio || 1;
    ctx.save();
    ctx.scale(dpr, dpr);

    // 绘制引用箭头（蓝色）
    ctx.strokeStyle = '#4472C4';
    ctx.fillStyle = '#4472C4';
    ctx.lineWidth = 1.5;
    arrows.precedents.forEach(arrow => {
        drawArrow(ctx, sheet, arrow.from, arrow.to);
    });

    // 绘制从属箭头（红色）
    ctx.strokeStyle = '#C44444';
    ctx.fillStyle = '#C44444';
    arrows.dependents.forEach(arrow => {
        drawArrow(ctx, sheet, arrow.from, arrow.to);
    });

    ctx.restore();
}

/**
 * 绘制单个箭头
 */
function drawArrow(ctx, sheet, from, to) {
    const fromInfo = sheet.getCellInViewInfo(from.r, from.c, false);
    const toInfo = sheet.getCellInViewInfo(to.r, to.c, false);

    if (!fromInfo.inView && !toInfo.inView) return;

    // 计算单元格中心点
    const x1 = fromInfo.x + fromInfo.w / 2;
    const y1 = fromInfo.y + fromInfo.h / 2;
    const x2 = toInfo.x + toInfo.w / 2;
    const y2 = toInfo.y + toInfo.h / 2;

    // 绘制线条
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();

    // 绘制起点圆点
    ctx.beginPath();
    ctx.arc(x1, y1, 4, 0, Math.PI * 2);
    ctx.fill();

    // 绘制箭头头部
    const angle = Math.atan2(y2 - y1, x2 - x1);
    const headLen = 10;
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - headLen * Math.cos(angle - Math.PI / 6), y2 - headLen * Math.sin(angle - Math.PI / 6));
    ctx.lineTo(x2 - headLen * Math.cos(angle + Math.PI / 6), y2 - headLen * Math.sin(angle + Math.PI / 6));
    ctx.closePath();
    ctx.fill();
}
