// 输入框相关监听

export function formulaBarFocus() {
    const oEditValue = this.formulaBar.value // 保留原始数据
    this._fbOriginalValue = oEditValue; // 供 Escape 取消使用
    if (!this.opCancel) {
        this.opCancel = this.SN.containerDom.querySelector('.sn-op-cancel')
        this.opConfirm = this.SN.containerDom.querySelector('.sn-op-confirm')
        this.opFun = this.SN.containerDom.querySelector('.sn-op-fun')
    }
    this.showCellInput(false, false);
    this.opCancel.style.color = 'red'
    this.opConfirm.style.color = 'green'
    this.opFun.style.color = 'blue'
    const inputFun = () => {
        this.input.innerText = this.formulaBar.value
    }
    this.formulaBar.addEventListener('input', inputFun)
    this.opCancel.addEventListener('mousedown', () => {
        this.input.innerText = oEditValue
    }, { once: true })
    this.formulaBar.addEventListener('blur', () => {
        this.updInputValue()
        this.formulaBar.removeEventListener('input', inputFun)
        this.opConfirm.style.color = ''
        this.opCancel.style.color = ''
        this.opFun.style.color = ''
    }, { once: true })
}

export function inputInput() {
    this.formulaBar.value = this.input.innerText
}

export function formulaBarKeyDown(event) {
    // Escape：取消编辑，恢复原始值
    if (event.key === 'Escape') {
        event.preventDefault();
        this.input.innerText = this._fbOriginalValue ?? '';
        this.formulaBar.blur();
        return;
    }
    if (event.key !== 'Enter') return;
    // Alt+Enter：插入换行（与 Excel 一致）
    if (event.altKey) {
        event.preventDefault();
        const s = this.formulaBar.selectionStart;
        const e = this.formulaBar.selectionEnd;
        const v = this.formulaBar.value;
        this.formulaBar.value = v.substring(0, s) + '\n' + v.substring(e);
        this.formulaBar.selectionStart = this.formulaBar.selectionEnd = s + 1;
        // 同步到单元格编辑器
        this.input.innerText = this.formulaBar.value;
        // 自动撑开高度适应内容
        const fb = this.formulaBar;
        fb.style.height = 'auto';
        fb.style.height = Math.min(fb.scrollHeight, 200) + 'px';
        return;
    }
    // Enter：提交
    event.preventDefault();
    this.formulaBar.blur();
}
