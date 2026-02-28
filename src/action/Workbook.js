// Workbook actions

export function fullScreen() {
    if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
    } else {
        document.exitFullscreen();
    }
}

export function moreSheet(dom) {
    let html = '';
    const ns = this.SN.namespace;
    this.SN.sheets.forEach((sheet) => {
        if (sheet.name === this.SN.activeSheet.name) {
            html += `<li class="sn-text-primary" data-ds="true">${sheet.name}</li>`;
            return;
        }
        if (sheet.hidden) return;
        html += `<li onmousedown="${ns}.activeSheet=${ns}.getSheet('${sheet.name}')" ontouchstart="${ns}.activeSheet=${ns}.getSheet('${sheet.name}')">${sheet.name}</li>`;
    });
    dom.nextElementSibling.innerHTML = html;
}

export function rename() {
    const promptText = this.SN.t('workbook.renamePrompt');
    const newName = prompt(promptText, this.SN.workbookName);
    if (newName) this.SN.workbookName = newName;
}

export function newWorkbook() {
    this.SN.IO._setData({
        version: '2.0',
        workbookName: this.SN.t('workbook.defaultName'),
        activeSheet: 'Sheet1',
        styleTable: { fonts: [], fills: [], borders: [], aligns: [] },
        sheets: [{ name: 'Sheet1', rows: [] }]
    });
}

export function confirmNewWorkbook() {
    const SN = this.SN;

    SN.Utils.modal({
        titleKey: 'workbook.modal.newWorkbook.title',
        title: SN.t('workbook.modal.newWorkbook.title'),
        contentKey: 'workbook.modal.newWorkbook.content',
        content: SN.t('workbook.modal.newWorkbook.content'),
        confirmKey: 'common.ok',
        confirmText: SN.t('common.ok'),
        cancelKey: 'common.cancel',
        cancelText: SN.t('common.cancel'),
        maxWidth: '400px'
    }).then((result) => {
        if (!result) return;
        this.newWorkbook();
    });
}

export function importFromUrl() {
    const t = (key) => this.SN.t(key);
    const html = `
        <input type="text" class="sn-url-input" placeholder="${t('workbook.importFromUrl.placeholder')}" style="width:100%;margin:2px 0 8px">
        <div style="color:#999;font-size:12px;">${t('workbook.importFromUrl.hint')}</div>
    `;
    this.SN.Utils.modal({
        titleKey: 'workbook.importFromUrl.title',
        title: t('workbook.importFromUrl.title'),
        content: html,
        confirmKey: 'workbook.importFromUrl.confirm',
        confirmText: t('workbook.importFromUrl.confirm'),
        cancelKey: 'common.cancel',
        cancelText: t('common.cancel'),
        maxWidth: '400px',
        onConfirm: () => {
            const url = this.SN.containerDom.querySelector('.sn-url-input')?.value?.trim();
            if (!url) return Promise.reject(t('workbook.importFromUrl.errors.urlRequired'));
            this.SN.IO.importFromUrl(url);
            return true;
        }
    });
}

export function showProperties() {
    const SN = this.SN;
    const props = SN.properties;
    const t = (key, params = {}) => SN.t(key, params);

    const formatDate = (isoStr) => {
        if (!isoStr) return '';
        const d = new Date(isoStr);
        const y = d.getFullYear();
        const m = d.getMonth() + 1;
        const day = d.getDate();
        const h = d.getHours();
        const min = d.getMinutes();
        return `${y}/${m}/${day} ${h}:${String(min).padStart(2, '0')}`;
    };

    const fileSize = props.fileSize
        ? t('action.workbook.properties.fileSize', { size: (props.fileSize / 1024).toFixed(1) })
        : t('action.workbook.properties.notAvailable');

    const inputStyle = 'width:100%;padding:4px 6px;border:none;border-bottom:1px solid #ddd;box-sizing:border-box;background:transparent;color:#0066cc;font-size:13px;outline:none;';
    const readonlyStyle = 'width:100%;padding:4px 6px;border:none;box-sizing:border-box;background:transparent;color:#666;font-size:13px;';
    const rowStyle = 'display:flex;align-items:flex-start;margin-bottom:6px;min-height:24px;';
    const labelStyle = 'width:100px;flex-shrink:0;color:#666;font-size:13px;padding-top:4px;';
    const sectionStyle = 'font-weight:bold;color:#666;font-size:13px;margin:16px 0 10px;';
    const personRowStyle = 'display:flex;align-items:center;margin-bottom:8px;';
    const avatarStyle = 'width:32px;height:32px;border-radius:50%;background:#e0e0e0;display:flex;align-items:center;justify-content:center;margin-right:10px;font-size:14px;color:#666;flex-shrink:0;';

    const html = `
        <div class="sn-scrollbar" style="max-height:480px;overflow-y:auto;padding:0 8px;">
            <div style="${sectionStyle}">${t('action.workbook.properties.section.properties')}</div>

            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.size')}</label>
                <span style="${readonlyStyle}">${fileSize}</span>
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.title')}</label>
                <input type="text" data-prop="title" value="${props.title || ''}" placeholder="${t('action.workbook.properties.placeholder.title')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.keywords')}</label>
                <input type="text" data-prop="keywords" value="${props.keywords || ''}" placeholder="${t('action.workbook.properties.placeholder.keywords')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.description')}</label>
                <input type="text" data-prop="description" value="${props.description || ''}" placeholder="${t('action.workbook.properties.placeholder.description')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.template')}</label>
                <span style="${readonlyStyle}">${props.template || ''}</span>
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.status')}</label>
                <input type="text" data-prop="contentStatus" value="${props.contentStatus || ''}" placeholder="${t('action.workbook.properties.placeholder.status')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.category')}</label>
                <input type="text" data-prop="category" value="${props.category || ''}" placeholder="${t('action.workbook.properties.placeholder.category')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.subject')}</label>
                <input type="text" data-prop="subject" value="${props.subject || ''}" placeholder="${t('action.workbook.properties.placeholder.subject')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.hyperlinkBase')}</label>
                <input type="text" data-prop="hyperlinkBase" value="${props.hyperlinkBase || ''}" placeholder="${t('action.workbook.properties.placeholder.hyperlinkBase')}" style="${inputStyle}">
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.company')}</label>
                <input type="text" data-prop="company" value="${props.company || ''}" placeholder="${t('action.workbook.properties.placeholder.company')}" style="${inputStyle}">
            </div>

            <div style="${sectionStyle}">${t('action.workbook.properties.section.dates')}</div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.lastModified')}</label>
                <span style="${readonlyStyle}">${formatDate(props.modified)}</span>
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.created')}</label>
                <span style="${readonlyStyle}">${formatDate(props.created)}</span>
            </div>
            <div style="${rowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.lastPrinted')}</label>
                <span style="${readonlyStyle}">${formatDate(props.lastPrinted) || ''}</span>
            </div>

            <div style="${sectionStyle}">${t('action.workbook.properties.section.people')}</div>
            <div style="${personRowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.manager')}</label>
                <div style="flex:1;display:flex;align-items:center;">
                    <div style="${avatarStyle}">${t('action.workbook.properties.avatar.manager')}</div>
                    <input type="text" data-prop="manager" value="${props.manager || ''}" placeholder="${t('action.workbook.properties.placeholder.manager')}" style="${inputStyle}flex:1;">
                </div>
            </div>
            <div style="${personRowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.author')}</label>
                <div style="flex:1;display:flex;align-items:center;">
                    <div style="${avatarStyle}">${(props.creator || t('action.workbook.properties.avatar.authorFallback'))[0].toUpperCase()}</div>
                    <input type="text" data-prop="creator" value="${props.creator || ''}" placeholder="${t('action.workbook.properties.placeholder.author')}" style="${inputStyle}flex:1;">
                </div>
            </div>
            <div style="${personRowStyle}">
                <label style="${labelStyle}">${t('action.workbook.properties.label.lastModifiedBy')}</label>
                <div style="flex:1;display:flex;align-items:center;">
                    <div style="${avatarStyle}">${(props.lastModifiedBy || t('action.workbook.properties.avatar.userFallback'))[0].toUpperCase()}</div>
                    <span style="${readonlyStyle}">${props.lastModifiedBy || ''}</span>
                </div>
            </div>
        </div>
    `;

    SN.Utils.modal({
        titleKey: 'action.workbook.modal.m003.title', title: t('action.workbook.modal.m003.title'),
        content: html,
        maxWidth: '420px',
        confirmKey: 'action.workbook.modal.m003.confirm', confirmText: t('action.workbook.modal.m003.confirm'),
        cancelKey: 'action.workbook.modal.m003.cancel', cancelText: t('action.workbook.modal.m003.cancel'),
        onConfirm: (bodyEl) => {
            const inputs = bodyEl.querySelectorAll('[data-prop]');
            inputs.forEach((input) => {
                const prop = input.dataset.prop;
                const val = input.value.trim();
                props[prop] = val;
            });

            SN.Utils.toast(SN.t('action.workbook.toast.propertiesSaved'));
            return true;
        }
    });
}
