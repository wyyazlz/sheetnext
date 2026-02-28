// 工作表保护模块

/**
 * 打开保护配置弹窗
 */
export function configProtection() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const p = sheet.protection;
    const t = (key, params = {}) => SN.t(key, params);

    // 保护选项配置
    const protectionOptions = [
        { key: 'selectLockedCells', labelKey: 'action.protection.content.sheet.options.selectLockedCells', default: true },
        { key: 'selectUnlockedCells', labelKey: 'action.protection.content.sheet.options.selectUnlockedCells', default: true },
        { key: 'formatCells', labelKey: 'action.protection.content.sheet.options.formatCells', default: false },
        { key: 'formatColumns', labelKey: 'action.protection.content.sheet.options.formatColumns', default: false },
        { key: 'formatRows', labelKey: 'action.protection.content.sheet.options.formatRows', default: false },
        { key: 'insertColumns', labelKey: 'action.protection.content.sheet.options.insertColumns', default: false },
        { key: 'insertRows', labelKey: 'action.protection.content.sheet.options.insertRows', default: false },
        { key: 'insertHyperlinks', labelKey: 'action.protection.content.sheet.options.insertHyperlinks', default: false },
        { key: 'deleteColumns', labelKey: 'action.protection.content.sheet.options.deleteColumns', default: false },
        { key: 'deleteRows', labelKey: 'action.protection.content.sheet.options.deleteRows', default: false },
        { key: 'sort', labelKey: 'action.protection.content.sheet.options.sort', default: false },
        { key: 'autoFilter', labelKey: 'action.protection.content.sheet.options.autoFilter', default: false },
        { key: 'pivotTables', labelKey: 'action.protection.content.sheet.options.pivotTables', default: false },
        { key: 'objects', labelKey: 'action.protection.content.sheet.options.objects', default: false },
    ];

    // 构建选项HTML（读取当前状态）
    const optionsHtml = protectionOptions.map(opt => {
        const checked = p[opt.key] ?? opt.default;
        return `
            <label class="sn-protect-option">
                <input type="checkbox" data-key="${opt.key}" ${checked ? 'checked' : ''}>
                <span>${t(opt.labelKey)}</span>
            </label>
        `;
    }).join('');

    const content = `
        <div class="sn-protect-sheet">
            <div class="sn-protect-row">
                <label>${t('action.protection.content.sheet.passwordLabel')}</label>
                <input type="password" class="sn-protect-password" placeholder="${t('action.protection.content.sheet.passwordPlaceholder')}">
            </div>
            <div class="sn-protect-label">${t('action.protection.content.sheet.allowAllUsers')}</div>
            <div class="sn-protect-options">
                ${optionsHtml}
            </div>
        </div>
    `;

    SN.Utils.modal({
        titleKey: 'action.protection.modal.m001.title', title: t('action.protection.modal.m001.title'),
        content,
        maxWidth: '400px',
        confirmKey: 'action.protection.modal.m001.confirm', confirmText: t('action.protection.modal.m001.confirm'),
        cancelKey: 'action.protection.modal.m001.cancel', cancelText: t('action.protection.modal.m001.cancel'),
        onConfirm: (bodyEl) => {
            const password = bodyEl.querySelector('.sn-protect-password').value;

            // 收集保护选项
            const options = { password };

            // 收集各选项状态
            bodyEl.querySelectorAll('.sn-protect-option input').forEach(input => {
                const key = input.dataset.key;
                options[key] = input.checked;
            });

            sheet.protection.enable(options);
            SN._r();
            SN.Layout?.refreshActiveButtons?.();
            SN.Utils.toast(SN.t('action.protection.toast.msg001'));
        }
    });
}

/**
 * 打开单元格保护配置弹窗（锁定/隐藏）
 */
export function configCellProtection() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    if (!sheet) return;
    const t = (key, params = {}) => SN.t(key, params);

    const getTargetAreas = () => {
        const { r, c } = sheet.activeCell || {};
        if (sheet.activeAreas?.length) return sheet.activeAreas;
        if (r === undefined || c === undefined) return [];
        return [{ s: { r, c }, e: { r, c } }];
    };

    const areas = getTargetAreas();
    let lockedTrue = false;
    let lockedFalse = false;
    let hiddenTrue = false;
    let hiddenFalse = false;
    areas.forEach((region) => {
        const range = typeof region === 'string' ? sheet.rangeStrToNum(region) : region;
        sheet.eachCells(range, (r, c) => {
            const cell = sheet.getCell(r, c);
            const locked = cell.protection.locked !== false;
            const hidden = !!cell.protection.hidden;
            if (locked) lockedTrue = true;
            else lockedFalse = true;
            if (hidden) hiddenTrue = true;
            else hiddenFalse = true;
        });
    });

    const lockedMixed = lockedTrue && lockedFalse;
    const hiddenMixed = hiddenTrue && hiddenFalse;
    const lockedChecked = lockedTrue && !lockedFalse;
    const hiddenChecked = hiddenTrue && !hiddenFalse;

    const content = `
        <div class="sn-protect-sheet sn-cell-protect">
            <div class="sn-protect-row">
                <label class="sn-protect-option">
                    <input type="checkbox" class="sn-cell-protect-lock" ${lockedChecked ? 'checked' : ''}>
                    <span>${t('action.protection.content.cell.locked')}</span>
                </label>
                <div class="sn-protect-label">${t('action.protection.content.cell.lockedDesc')}</div>
            </div>
            <div class="sn-protect-row">
                <label class="sn-protect-option">
                    <input type="checkbox" class="sn-cell-protect-hide" ${hiddenChecked ? 'checked' : ''}>
                    <span>${t('action.protection.content.cell.hiddenFormula')}</span>
                </label>
                <div class="sn-protect-label">${t('action.protection.content.cell.hiddenFormulaDesc')}</div>
            </div>
            <div class="sn-protect-label">${t('action.protection.content.cell.tip')}</div>
        </div>
    `;

    SN.Utils.modal({
        titleKey: 'action.protection.modal.m002.title', title: t('action.protection.modal.m002.title'),
        content,
        maxWidth: '360px',
        confirmKey: 'action.protection.modal.m002.confirm', confirmText: t('action.protection.modal.m002.confirm'),
        cancelKey: 'action.protection.modal.m002.cancel', cancelText: t('action.protection.modal.m002.cancel'),
        onOpen: (bodyEl) => {
            const lockCheckbox = bodyEl.querySelector('.sn-cell-protect-lock');
            const hideCheckbox = bodyEl.querySelector('.sn-cell-protect-hide');
            if (lockCheckbox) lockCheckbox.indeterminate = lockedMixed;
            if (hideCheckbox) hideCheckbox.indeterminate = hiddenMixed;
        },
        onConfirm: (bodyEl) => {
            const lockCheckbox = bodyEl.querySelector('.sn-cell-protect-lock');
            const hideCheckbox = bodyEl.querySelector('.sn-cell-protect-hide');
            const options = {};

            if (lockCheckbox && !lockCheckbox.indeterminate) options.locked = lockCheckbox.checked;
            if (hideCheckbox && !hideCheckbox.indeterminate) options.hidden = hideCheckbox.checked;

            if (Object.keys(options).length === 0) return;

            const targetAreas = getTargetAreas();
            if (!targetAreas.length) return;

            sheet.protection.setCellProtection(targetAreas, options);
            SN._r();
            SN.Layout?.refreshActiveButtons?.();
            SN.Utils.toast(SN.t('action.protection.toast.msg002'));
        }
    });
}

/**
 * 取消保护工作表
 */
export function cancelProtection() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const p = sheet.protection;
    const t = (key, params = {}) => SN.t(key, params);

    // 未保护时提示
    if (!p.enabled) {
        return SN.Utils.toast(SN.t('action.protection.toast.msg003'));
    }

    const opts = p.getOptions();

    // 如果设置了密码
    if (opts?.passwordHash || opts?.hashValue) {
        const content = `
            <div class="sn-protect-sheet">
                <div class="sn-protect-row">
                    <label>${t('action.protection.content.unprotect.passwordLabel')}</label>
                    <input type="password" class="sn-unprotect-password" placeholder="${t('action.protection.content.unprotect.passwordPlaceholder')}">
                </div>
            </div>
        `;

        SN.Utils.modal({
            titleKey: 'action.protection.modal.m003.title', title: t('action.protection.modal.m003.title'),
            content,
            maxWidth: '320px',
            confirmKey: 'action.protection.modal.m003.confirm', confirmText: t('action.protection.modal.m003.confirm'),
            cancelKey: 'action.protection.modal.m003.cancel', cancelText: t('action.protection.modal.m003.cancel'),
            onOpen: (bodyEl) => {
                bodyEl.querySelector('.sn-unprotect-password').focus();
            },
            onConfirm: (bodyEl) => {
                const inputPassword = bodyEl.querySelector('.sn-unprotect-password').value;
                const result = p.disable(inputPassword);
                if (!result.ok) throw result.message;
                SN._r();
                SN.Layout?.refreshActiveButtons?.();
                SN.Utils.toast(SN.t('action.protection.toast.msg004'));
            }
        });
    } else {
        // 没有密码，直接取消保护
        p.disable();
        SN._r();
        SN.Layout?.refreshActiveButtons?.();
        SN.Utils.toast(SN.t('action.protection.toast.msg004'));
    }
}

/**
 * 打开保护工作表弹窗（或取消保护）
 */
export function protectSheet() {
    const sheet = this.SN.activeSheet;
    if (sheet.protection.enabled) {
        return cancelProtection.call(this);
    }
    return configProtection.call(this);
}

/**
 * 取消保护工作表
 */
export function unprotectSheet() {
    return cancelProtection.call(this);
}

/**
 * 设置/重置权限
 */
export function setPermission(type, value) {
    const SN = this.SN;
    SN.Utils.toast(SN.t('action.protection.toast.permissionUpdated', { type, value }));
}

/**
 * 重置所有权限
 */
export function resetPermissions() {
    const SN = this.SN;
    SN.activeSheet.protection.disable();
    SN._r();
    SN.Layout?.refreshActiveButtons?.();
    SN.Utils.toast(SN.t('action.protection.toast.msg005'));
}
