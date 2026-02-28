/**
 * 查找替换定位对话框
 * 复刻 Excel 的查找、替换、定位功能
 */

export default class FindReplaceDialog {
    constructor(SN) {
        this.SN = SN;
        this.overlay = null;
        this.activeTab = 'find'; // find | replace | goto
        this.findResults = [];   // 查找结果
        this.currentIndex = -1;  // 当前结果索引
        this.lastSearchText = '';
        this.options = {
            matchCase: false,      // 区分大小写
            matchCell: false,      // 单元格匹配
            searchIn: 'values',    // values | formulas
            searchBy: 'rows',      // rows | columns
            searchScope: 'sheet'   // sheet | workbook
        };
    }

    /**
     * 显示对话框
     * @param {string} tab - 初始激活的标签: find | replace | goto
     */
    show(tab = 'find') {
        this.activeTab = tab;
        this._createDialog();
        this._show();
    }

    /**
     * 创建对话框
     */
    _createDialog() {
        if (this.overlay) {
            if (this._escHandler) {
                document.removeEventListener('keydown', this._escHandler);
                this._escHandler = null;
            }
            this.overlay.remove();
            this.overlay = null;
        }

        const overlay = document.createElement('div');
        overlay.className = 'sn-modal-overlay sn-find-dialog-overlay';
        const html = `
            <div class="sn-modal sn-find-dialog">
                <div class="sn-modal-header">
                    <div class="sn-find-tabs">
                        <button class="sn-find-tab active" data-tab="find">${this._t('findReplace.tab.find', 'Find')}</button>
                        <button class="sn-find-tab" data-tab="replace">${this._t('findReplace.tab.replace', 'Replace')}</button>
                        <button class="sn-find-tab" data-tab="goto">${this._t('findReplace.tab.goto', 'Go To')}</button>
                    </div>
                    <button class="sn-modal-close" title="${this._t('findReplace.close', 'Close')}"></button>
                </div>
                <div class="sn-modal-body">
                    <!-- 查找面板 -->
                    <div class="sn-find-panel active" data-panel="find">
                        ${this._renderFindPanel()}
                    </div>
                    <!-- 替换面板 -->
                    <div class="sn-find-panel" data-panel="replace">
                        ${this._renderReplacePanel()}
                    </div>
                    <!-- 定位面板 -->
                    <div class="sn-find-panel" data-panel="goto">
                        ${this._renderGotoPanel()}
                    </div>
                </div>
            </div>
        `;
        overlay.innerHTML = html;

        document.body.appendChild(overlay);
        this.overlay = overlay;

        this._bindEvents();
        this._switchTab(this.activeTab);
    }

    /**
     * 渲染查找面板
     */
    _renderFindPanel() {
        return `
            <div class="sn-find-form">
                <div class="sn-find-row">
                    <label>${this._t('findReplace.findWhat', 'Find what:')}</label>
                    <input type="text" class="sn-find-input" data-field="find-text" placeholder="${this._t('findReplace.placeholder.searchText', 'Enter search text')}">
                </div>
                <div class="sn-find-options">
                    <details class="sn-find-options-toggle">
                        <summary>${this._t('findReplace.options', 'Options')}</summary>
                        <div class="sn-find-options-content">
                            <div class="sn-find-options-row">
                                <label class="sn-find-checkbox">
                                    <input type="checkbox" data-field="find-match-case">
                                    <span>${this._t('findReplace.matchCase', 'Match case')}</span>
                                </label>
                                <label class="sn-find-checkbox">
                                    <input type="checkbox" data-field="find-match-cell">
                                    <span>${this._t('findReplace.matchEntireCell', 'Match entire cell contents')}</span>
                                </label>
                            </div>
                            <div class="sn-find-options-row">
                                <label>${this._t('findReplace.within', 'Within:')}</label>
                                <select data-field="find-scope">
                                    <option value="sheet">${this._t('findReplace.scope.sheet', 'Sheet')}</option>
                                    <option value="workbook">${this._t('findReplace.scope.workbook', 'Workbook')}</option>
                                </select>
                                <label>${this._t('findReplace.search', 'Search:')}</label>
                                <select data-field="find-by">
                                    <option value="rows">${this._t('findReplace.searchBy.rows', 'By Rows')}</option>
                                    <option value="columns">${this._t('findReplace.searchBy.columns', 'By Columns')}</option>
                                </select>
                            </div>
                            <div class="sn-find-options-row">
                                <label>${this._t('findReplace.lookIn', 'Look in:')}</label>
                                <select data-field="find-in">
                                    <option value="values">${this._t('findReplace.lookInValues', 'Values')}</option>
                                    <option value="formulas">${this._t('findReplace.lookInFormulas', 'Formulas')}</option>
                                </select>
                            </div>
                        </div>
                    </details>
                </div>
                <div class="sn-find-status" data-field="find-status"></div>
            </div>
            <div class="sn-find-actions">
                <button class="sn-modal-btn" data-field="find-all">${this._t('findReplace.findAll', 'Find All')}</button>
                <button class="sn-modal-btn sn-modal-btn-confirm" data-field="find-next">${this._t('findReplace.findNext', 'Find Next')}</button>
            </div>
            <div class="sn-find-results" data-field="find-results"></div>
        `;
    }

    /**
     * 渲染替换面板
     */
    _renderReplacePanel() {
        return `
            <div class="sn-find-form">
                <div class="sn-find-row">
                    <label>${this._t('findReplace.findWhat', 'Find what:')}</label>
                    <input type="text" class="sn-find-input" data-field="replace-find" placeholder="${this._t('findReplace.placeholder.searchText', 'Enter search text')}">
                </div>
                <div class="sn-find-row">
                    <label>${this._t('findReplace.replaceWith', 'Replace with:')}</label>
                    <input type="text" class="sn-find-input" data-field="replace-with" placeholder="${this._t('findReplace.placeholder.replaceText', 'Enter replacement text')}">
                </div>
                <div class="sn-find-options">
                    <details class="sn-find-options-toggle">
                        <summary>${this._t('findReplace.options', 'Options')}</summary>
                        <div class="sn-find-options-content">
                            <div class="sn-find-options-row">
                                <label class="sn-find-checkbox">
                                    <input type="checkbox" data-field="replace-match-case">
                                    <span>${this._t('findReplace.matchCase', 'Match case')}</span>
                                </label>
                                <label class="sn-find-checkbox">
                                    <input type="checkbox" data-field="replace-match-cell">
                                    <span>${this._t('findReplace.matchEntireCell', 'Match entire cell contents')}</span>
                                </label>
                            </div>
                            <div class="sn-find-options-row">
                                <label>${this._t('findReplace.within', 'Within:')}</label>
                                <select data-field="replace-scope">
                                    <option value="sheet">${this._t('findReplace.scope.sheet', 'Sheet')}</option>
                                    <option value="workbook">${this._t('findReplace.scope.workbook', 'Workbook')}</option>
                                </select>
                                <label>${this._t('findReplace.lookIn', 'Look in:')}</label>
                                <select data-field="replace-in">
                                    <option value="values">${this._t('findReplace.lookInValues', 'Values')}</option>
                                    <option value="formulas">${this._t('findReplace.lookInFormulas', 'Formulas')}</option>
                                </select>
                            </div>
                        </div>
                    </details>
                </div>
                <div class="sn-find-status" data-field="replace-status"></div>
            </div>
            <div class="sn-find-actions">
                <button class="sn-modal-btn" data-field="replace-all">${this._t('findReplace.replaceAll', 'Replace All')}</button>
                <button class="sn-modal-btn" data-field="replace-one">${this._t('findReplace.tab.replace', 'Replace')}</button>
                <button class="sn-modal-btn sn-modal-btn-confirm" data-field="replace-find-next">${this._t('findReplace.findNext', 'Find Next')}</button>
            </div>
        `;
    }

    /**
     * 渲染定位面板
     */
    _renderGotoPanel() {
        return `
            <div class="sn-find-form">
                <div class="sn-find-row">
                    <label>${this._t('findReplace.reference', 'Reference:')}</label>
                    <input type="text" class="sn-find-input" data-field="goto-ref" placeholder="${this._t('findReplace.referencePlaceholder', 'e.g. A1, A1:B10, Sheet2!A1')}">
                </div>
                <div class="sn-goto-special">
                    <div class="sn-goto-special-title">${this._t('findReplace.gotoSpecialTitle', 'Go To Special:')}</div>
                    <div class="sn-goto-special-grid">
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="comments">
                            <span>${this._t('findReplace.special.comments', 'Comments')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="formulas">
                            <span>${this._t('findReplace.special.formulas', 'Formulas')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="constants">
                            <span>${this._t('findReplace.special.constants', 'Constants')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="blanks">
                            <span>${this._t('findReplace.special.blanks', 'Blanks')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="errors">
                            <span>${this._t('findReplace.special.errors', 'Errors')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="differences">
                            <span>${this._t('findReplace.special.rowDifferences', 'Row Differences')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="colDifferences">
                            <span>${this._t('findReplace.special.columnDifferences', 'Column Differences')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="precedents">
                            <span>${this._t('findReplace.special.precedents', 'Precedents')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="dependents">
                            <span>${this._t('findReplace.special.dependents', 'Dependents')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="lastCell">
                            <span>${this._t('findReplace.special.lastCell', 'Last Cell')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="visible">
                            <span>${this._t('findReplace.special.visibleCells', 'Visible Cells')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="conditional">
                            <span>${this._t('findReplace.special.conditionalFormats', 'Conditional Formats')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="validation">
                            <span>${this._t('findReplace.special.dataValidation', 'Data Validation')}</span>
                        </label>
                        <label class="sn-goto-option">
                            <input type="radio" name="goto-special" value="merged">
                            <span>${this._t('findReplace.special.mergedCells', 'Merged Cells')}</span>
                        </label>
                    </div>
                </div>
            </div>
            <div class="sn-find-actions">
                <button class="sn-modal-btn" data-field="goto-special-btn">${this._t('findReplace.gotoSpecial', 'Go To Special')}</button>
                <button class="sn-modal-btn sn-modal-btn-confirm" data-field="goto-btn">${this._t('findReplace.tab.goto', 'Go To')}</button>
            </div>
        `;
    }

    /**
     * 绑定事件
     */
    _bindEvents() {
        const overlay = this.overlay;

        // 关闭按钮
        overlay.querySelector('.sn-modal-close').addEventListener('click', () => this.hide());

        // Tab 切换
        overlay.querySelectorAll('.sn-find-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                this._switchTab(tab.dataset.tab);
            });
        });

        // 查找面板事件
        overlay.querySelector('[data-field="find-next"]')?.addEventListener('click', () => this._findNext());
        overlay.querySelector('[data-field="find-all"]')?.addEventListener('click', () => this._findAll());
        overlay.querySelector('[data-field="find-text"]')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') this._findNext();
        });

        // 替换面板事件
        overlay.querySelector('[data-field="replace-find-next"]')?.addEventListener('click', () => this._findNextReplace());
        overlay.querySelector('[data-field="replace-one"]')?.addEventListener('click', () => this._replaceOne());
        overlay.querySelector('[data-field="replace-all"]')?.addEventListener('click', () => this._replaceAll());
        overlay.querySelector('[data-field="replace-find"]')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') this._findNextReplace();
        });

        // 定位面板事件
        overlay.querySelector('[data-field="goto-btn"]')?.addEventListener('click', () => this._goto());
        overlay.querySelector('[data-field="goto-special-btn"]')?.addEventListener('click', () => this._gotoSpecial());
        overlay.querySelector('[data-field="goto-ref"]')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') this._goto();
        });

        // ESC 关闭
        this._escHandler = (e) => {
            if (e.key === 'Escape') this.hide();
        };
        document.addEventListener('keydown', this._escHandler);
    }

    /**
     * 切换标签
     */
    _switchTab(tab) {
        this.activeTab = tab;
        const overlay = this.overlay;

        // 更新 Tab 按钮状态
        overlay.querySelectorAll('.sn-find-tab').forEach(t => {
            t.classList.toggle('active', t.dataset.tab === tab);
        });

        // 更新面板显示
        overlay.querySelectorAll('.sn-find-panel').forEach(p => {
            p.classList.toggle('active', p.dataset.panel === tab);
        });

        // 自动聚焦输入框
        setTimeout(() => {
            const input = overlay.querySelector(`.sn-find-panel.active input[type="text"]`);
            if (input) input.focus();
        }, 50);
    }

    /**
     * 获取查找选项
     */
    _getOptions(prefix = 'find') {
        const overlay = this.overlay;
        return {
            matchCase: overlay.querySelector(`[data-field="${prefix}-match-case"]`)?.checked || false,
            matchCell: overlay.querySelector(`[data-field="${prefix}-match-cell"]`)?.checked || false,
            searchIn: overlay.querySelector(`[data-field="${prefix}-in"]`)?.value || 'values',
            searchBy: overlay.querySelector(`[data-field="${prefix}-by"]`)?.value || 'rows',
            searchScope: overlay.querySelector(`[data-field="${prefix}-scope"]`)?.value || 'sheet'
        };
    }

    /**
     * 查找下一个
     */
    _findNext() {
        const searchText = this.overlay.querySelector('[data-field="find-text"]').value;
        if (!searchText) {
            this._setStatus('find', 'findReplace.status.enterSearchText', 'Please enter search text');
            return;
        }

        const options = this._getOptions('find');

        // 如果搜索文本或选项改变，重新搜索
        if (searchText !== this.lastSearchText || this.findResults.length === 0) {
            this._performSearch(searchText, options);
            this.lastSearchText = searchText;
        }

        if (this.findResults.length === 0) {
            this._setStatus('find', 'findReplace.status.notFound', '"{text}" not found', { text: searchText });
            return;
        }

        // 移动到下一个结果
        this.currentIndex = (this.currentIndex + 1) % this.findResults.length;
        this._gotoResult(this.findResults[this.currentIndex]);
        this._setStatus('find', 'findReplace.status.progress', '{current} / {total}', {
            current: this.currentIndex + 1,
            total: this.findResults.length
        });
    }

    /**
     * 查找全部
     */
    _findAll() {
        const searchText = this.overlay.querySelector('[data-field="find-text"]').value;
        if (!searchText) {
            this._setStatus('find', 'findReplace.status.enterSearchText', 'Please enter search text');
            return;
        }

        const options = this._getOptions('find');
        this._performSearch(searchText, options);
        this.lastSearchText = searchText;

        if (this.findResults.length === 0) {
            this._setStatus('find', 'findReplace.status.notFound', '"{text}" not found', { text: searchText });
            this._renderResults([]);
            return;
        }

        this._setStatus('find', 'findReplace.status.foundCount', 'Found {count} cells', { count: this.findResults.length });
        this._renderResults(this.findResults);
    }

    /**
     * 执行搜索
     */
    _performSearch(searchText, options) {
        this.findResults = [];
        this.currentIndex = -1;

        const sheets = options.searchScope === 'workbook'
            ? this.SN.sheets
            : [this.SN.activeSheet];

        const searchLower = options.matchCase ? searchText : searchText.toLowerCase();

        for (const sheet of sheets) {
            if (!sheet.initialized) sheet._init();

            const iterateByRows = options.searchBy === 'rows';

            if (iterateByRows) {
                for (let r = 0; r < sheet.rowCount; r++) {
                    for (let c = 0; c < sheet.colCount; c++) {
                        this._checkCell(sheet, r, c, searchLower, options);
                    }
                }
            } else {
                for (let c = 0; c < sheet.colCount; c++) {
                    for (let r = 0; r < sheet.rowCount; r++) {
                        this._checkCell(sheet, r, c, searchLower, options);
                    }
                }
            }
        }
    }

    /**
     * 检查单元格是否匹配
     */
    _checkCell(sheet, r, c, searchText, options) {
        const cell = sheet.getCell(r, c);
        if (!cell) return;

        let value = options.searchIn === 'formulas' && cell.isFormula
            ? cell.editVal
            : String(cell.showVal ?? cell.editVal ?? '');

        if (!options.matchCase) {
            value = value.toLowerCase();
        }

        let matched = false;
        if (options.matchCell) {
            matched = value === searchText;
        } else {
            matched = value.includes(searchText);
        }

        if (matched) {
            this.findResults.push({
                sheet: sheet,
                sheetName: sheet.name,
                r: r,
                c: c,
                value: String(cell.showVal ?? cell.editVal ?? ''),
                ref: `${this.SN.Utils.numToChar(c)}${r + 1}`
            });
        }
    }

    /**
     * 跳转到结果
     */
    _gotoResult(result) {
        if (result.sheet !== this.SN.activeSheet) {
            this.SN.activeSheet = result.sheet;
        }

        const sheet = this.SN.activeSheet;

        // 设置活动单元格和选区
        sheet.activeCell = { r: result.r, c: result.c };
        sheet.activeAreas = [{ s: { r: result.r, c: result.c }, e: { r: result.r, c: result.c } }];

        // 滚动视图到目标位置
        sheet.vi.rowArr[0] = Math.max(0, result.r - 2);
        sheet.vi.colArr[0] = Math.max(0, result.c - 1);

        this.SN._r();
    }

    /**
     * 渲染结果列表
     */
    _renderResults(results) {
        const container = this.overlay.querySelector('[data-field="find-results"]');
        if (!container) return;

        if (results.length === 0) {
            container.innerHTML = '';
            container.classList.remove('has-results');
            return;
        }

        container.classList.add('has-results');
        container.innerHTML = `
            <div class="sn-find-results-header">
                <span>${this._t('findReplace.result.sheet', 'Sheet')}</span>
                <span>${this._t('findReplace.result.cell', 'Cell')}</span>
                <span>${this._t('findReplace.result.value', 'Value')}</span>
            </div>
            <div class="sn-find-results-list">
                ${results.map((r, i) => `
                    <div class="sn-find-result-item" data-index="${i}">
                        <span>${r.sheetName}</span>
                        <span>${r.ref}</span>
                        <span title="${this._escapeHtml(r.value)}">${this._escapeHtml(r.value)}</span>
                    </div>
                `).join('')}
            </div>
        `;

        // 绑定点击事件
        container.querySelectorAll('.sn-find-result-item').forEach(item => {
            item.addEventListener('click', () => {
                const index = parseInt(item.dataset.index);
                this.currentIndex = index;
                this._gotoResult(this.findResults[index]);

                // 高亮选中项
                container.querySelectorAll('.sn-find-result-item').forEach(i =>
                    i.classList.toggle('active', i === item)
                );
            });
        });
    }

    /**
     * 替换面板 - 查找下一个
     */
    _findNextReplace() {
        const searchText = this.overlay.querySelector('[data-field="replace-find"]').value;
        if (!searchText) {
            this._setStatus('replace', 'findReplace.status.enterSearchText', 'Please enter search text');
            return;
        }

        const options = this._getOptions('replace');

        if (searchText !== this.lastSearchText || this.findResults.length === 0) {
            this._performSearch(searchText, options);
            this.lastSearchText = searchText;
        }

        if (this.findResults.length === 0) {
            this._setStatus('replace', 'findReplace.status.notFound', '"{text}" not found', { text: searchText });
            return;
        }

        this.currentIndex = (this.currentIndex + 1) % this.findResults.length;
        this._gotoResult(this.findResults[this.currentIndex]);
        this._setStatus('replace', 'findReplace.status.progress', '{current} / {total}', {
            current: this.currentIndex + 1,
            total: this.findResults.length
        });
    }

    /**
     * 替换当前
     */
    _replaceOne() {
        if (this.findResults.length === 0 || this.currentIndex < 0) {
            this._findNextReplace();
            return;
        }

        const replaceWith = this.overlay.querySelector('[data-field="replace-with"]').value;
        const result = this.findResults[this.currentIndex];
        const searchText = this.overlay.querySelector('[data-field="replace-find"]').value;
        const options = this._getOptions('replace');

        this._replaceCell(result, searchText, replaceWith, options);

        // 从结果中移除已替换的
        this.findResults.splice(this.currentIndex, 1);

        if (this.findResults.length === 0) {
            this._setStatus('replace', 'findReplace.status.replaceCompleted', 'Replace completed');
            this.currentIndex = -1;
            this.SN._r();
            return;
        }

        // 调整索引并跳转到下一个
        if (this.currentIndex >= this.findResults.length) {
            this.currentIndex = 0;
        }
        this._gotoResult(this.findResults[this.currentIndex]);
        this._setStatus('replace', 'findReplace.status.progress', '{current} / {total}', {
            current: this.currentIndex + 1,
            total: this.findResults.length
        });
    }

    /**
     * 全部替换
     */
    _replaceAll() {
        const searchText = this.overlay.querySelector('[data-field="replace-find"]').value;
        const replaceWith = this.overlay.querySelector('[data-field="replace-with"]').value;

        if (!searchText) {
            this._setStatus('replace', 'findReplace.status.enterSearchText', 'Please enter search text');
            return;
        }

        const options = this._getOptions('replace');
        this._performSearch(searchText, options);

        if (this.findResults.length === 0) {
            this._setStatus('replace', 'findReplace.status.notFound', '"{text}" not found', { text: searchText });
            return;
        }

        const count = this.findResults.length;

        // 批量替换
        for (const result of this.findResults) {
            this._replaceCell(result, searchText, replaceWith, options);
        }

        this.findResults = [];
        this.currentIndex = -1;
        this.lastSearchText = '';

        this.SN._r();
        this._setStatus('replace', 'findReplace.status.replacedCount', 'Replaced {count} occurrence(s)', { count });
    }

    /**
     * 替换单元格内容
     */
    _replaceCell(result, searchText, replaceWith, options) {
        const cell = result.sheet.getCell(result.r, result.c);
        if (!cell) return;

        let value = String(cell.editVal ?? '');

        if (options.matchCell) {
            // 整单元格替换
            value = replaceWith;
        } else {
            // 部分替换
            if (options.matchCase) {
                value = value.split(searchText).join(replaceWith);
            } else {
                const regex = new RegExp(this._escapeRegex(searchText), 'gi');
                value = value.replace(regex, replaceWith);
            }
        }

        // 直接设置单元格值
        cell.editVal = value;
        // 清除缓存的显示值
        cell._showVal = undefined;
        cell._calcVal = undefined;
    }

    /**
     * 定位
     */
    _goto() {
        const refInput = this.overlay.querySelector('[data-field="goto-ref"]');
        const ref = refInput.value.trim();

        if (!ref) {
            this.SN.Utils.toast(this._t('findReplace.toast.enterReference', 'Please enter a reference'));
            return;
        }

        try {
            // 解析引用
            const parsed = this._parseReference(ref);

            if (parsed.sheetName && parsed.sheetName !== this.SN.activeSheet.name) {
                const sheet = this.SN.getSheet(parsed.sheetName);
                if (!sheet) {
                    this.SN.Utils.toast(this._t('findReplace.toast.sheetNotFound', 'Worksheet "{name}" not found', { name: parsed.sheetName }));
                    return;
                }
                this.SN.activeSheet = sheet;
            }

            const sheet = this.SN.activeSheet;

            if (parsed.range) {
                // 选择区域
                sheet.activeCell = { r: parsed.range.s.r, c: parsed.range.s.c };
                sheet.activeAreas = [parsed.range];
                sheet.vi.rowArr[0] = Math.max(0, parsed.range.s.r - 2);
                sheet.vi.colArr[0] = Math.max(0, parsed.range.s.c - 1);
            } else {
                sheet.activeCell = { r: parsed.r, c: parsed.c };
                sheet.activeAreas = [{ s: { r: parsed.r, c: parsed.c }, e: { r: parsed.r, c: parsed.c } }];
                sheet.vi.rowArr[0] = Math.max(0, parsed.r - 2);
                sheet.vi.colArr[0] = Math.max(0, parsed.c - 1);
            }

            this.SN._r();
            this.hide();
        } catch (e) {
            this.SN.Utils.toast(this._t('findReplace.toast.invalidReference', 'Invalid reference format'));
        }
    }

    /**
     * 定位条件
     */
    _gotoSpecial() {
        const selected = this.overlay.querySelector('input[name="goto-special"]:checked');
        if (!selected) {
            this.SN.Utils.toast(this._t('findReplace.toast.selectGoToCondition', 'Please select a go-to condition'));
            return;
        }

        const type = selected.value;
        const sheet = this.SN.activeSheet;
        const results = [];

        // 获取搜索范围
        const areas = sheet.activeAreas;
        const searchArea = areas.length > 0 && (areas[0].e.r > areas[0].s.r || areas[0].e.c > areas[0].s.c)
            ? areas[0]
            : { s: { r: 0, c: 0 }, e: { r: sheet.rowCount - 1, c: sheet.colCount - 1 } };

        for (let r = searchArea.s.r; r <= searchArea.e.r; r++) {
            for (let c = searchArea.s.c; c <= searchArea.e.c; c++) {
                if (this._matchSpecial(sheet, r, c, type)) {
                    results.push({ r, c });
                }
            }
        }

        if (results.length === 0) {
            this.SN.Utils.toast(this._t('findReplace.toast.noMatchingCells', 'No matching cells found'));
            return;
        }

        // 选中所有匹配的单元格
        sheet.activeCell = results[0];
        sheet.activeAreas = results.map(p => ({
            s: { r: p.r, c: p.c },
            e: { r: p.r, c: p.c }
        }));

        this.SN._r();
        this.SN.Utils.toast(this._t('findReplace.toast.selectedCount', 'Selected {count} cells', { count: results.length }));
        this.hide();
    }

    /**
     * 检查单元格是否匹配定位条件
     */
    _matchSpecial(sheet, r, c, type) {
        const cell = sheet.getCell(r, c);

        switch (type) {
            case 'formulas':
                return cell.isFormula;
            case 'constants':
                return !cell.isFormula && cell.editVal !== undefined && cell.editVal !== null && cell.editVal !== '';
            case 'blanks':
                return cell.editVal === undefined || cell.editVal === null || cell.editVal === '';
            case 'errors':
                return cell.isFormula && String(cell.calcVal).startsWith('#');
            case 'comments':
                return cell.Comment && cell.Comment.length > 0;
            case 'conditional':
                return sheet.CF?.rules?.some(rule =>
                    rule.sqref?.split(' ').some(rangeStr => {
                        const range = sheet.rangeStrToNum(rangeStr);
                        return r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c;
                    })
                );
            case 'validation':
                return cell.dataValidation !== undefined;
            case 'merged':
                return sheet.merges?.some(m =>
                    r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c
                );
            case 'lastCell':
                // 找到最后使用的单元格
                return false; // 需要特殊处理
            case 'visible':
                return !sheet.rows[r]?.hidden && !sheet.cols[c]?.hidden;
            case 'precedents':
            case 'dependents':
            case 'differences':
            case 'colDifferences':
                // 这些需要更复杂的逻辑
                return false;
            default:
                return false;
        }
    }

    /**
     * 解析引用
     */
    _parseReference(ref) {
        let sheetName = null;
        let cellRef = ref;

        // 检查是否包含工作表名称
        if (ref.includes('!')) {
            const parts = ref.split('!');
            sheetName = parts[0].replace(/^'|'$/g, '');
            cellRef = parts[1];
        }

        // 检查是否是范围
        if (cellRef.includes(':')) {
            const [start, end] = cellRef.split(':');
            const s = this._parseCellRef(start);
            const e = this._parseCellRef(end);
            return { sheetName, range: { s, e } };
        }

        const pos = this._parseCellRef(cellRef);
        return { sheetName, r: pos.r, c: pos.c };
    }

    /**
     * 解析单元格引用
     */
    _parseCellRef(ref) {
        const match = ref.match(/^([A-Za-z]+)(\d+)$/);
        if (!match) throw new Error(this._t('findReplace.error.invalidCellReference', 'Invalid cell reference'));

        const col = this.SN.Utils.charToNum(match[1].toUpperCase());
        const row = parseInt(match[2]) - 1;

        return { r: row, c: col };
    }

    /**
     * 设置状态信息
     */
    _setStatus(panel, key, fallback, params = {}) {
        const statusEl = this.overlay.querySelector(`[data-field="${panel}-status"]`);
        if (statusEl) {
            statusEl.textContent = this._t(key, fallback, params);
        }
    }

    /** @param {string} key @param {string} _fallback @param {Record<string, any>} [params={}] @returns {string} */
    _t(key, _fallback, params = {}) {
        return this.SN.t(key, params);
    }

    /**
     * HTML 转义
     */
    _escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    /**
     * 正则转义
     */
    _escapeRegex(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * 更新内容
     */
    _updateContent() {
        // 可以在这里更新面板内容
    }

    /**
     * 显示对话框
     */
    _show() {
        // 重置拖动位置
        const modal = this.overlay.querySelector('.sn-modal');
        if (modal) modal.style.transform = '';

        requestAnimationFrame(() => {
            this.overlay.classList.add('active');

            // 聚焦到输入框
            const input = this.overlay.querySelector('.sn-find-panel.active input[type="text"]');
            if (input) {
                input.focus();
                input.select();
            }
        });
    }

    /**
     * 隐藏对话框
     */
    hide() {
        if (!this.overlay) return;

        this.overlay.classList.add('closing');
        setTimeout(() => {
            this.overlay.classList.remove('active', 'closing');
        }, 200);

        if (this._escHandler) {
            document.removeEventListener('keydown', this._escHandler);
            this._escHandler = null;
        }
    }

    /**
     * 销毁对话框
     */
    destroy() {
        this.hide();
        if (this.overlay) {
            this.overlay.remove();
            this.overlay = null;
        }
    }
}
