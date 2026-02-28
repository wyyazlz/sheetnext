export default class UndoRedo {
    /** @param {import('../Workbook/Workbook.js').default} SN */
    constructor(SN) {
        /** @type {number} */
        this.maxStack = 50;

        this._SN = SN;
        this._undoStack = [];
        this._redoStack = [];
        this._currentTx = null;
        this._manualMode = false;
        this._manualDepth = 0; // Manual transaction nesting level.
        this._microTaskPending = false;
    }

    /** @type {boolean} */
    get canUndo() {
        return this._undoStack.length > 0 || (!this._manualMode && this._currentTx?.actions?.length > 0);
    }

    /** @type {boolean} */
    get canRedo() {
        return this._redoStack.length > 0;
    }

    /** @param {{undo: Function, redo: Function, activeCell?: {r:number,c:number}}} action @returns {void} Add undo action into current transaction. */
    add(action) {
        if (this._currentTx) {
            this._currentTx.actions.push(action);
        } else {
            const viewInfo = this._captureView();
            if (action.activeCell) {
                viewInfo.activeCell = { r: action.activeCell.r, c: action.activeCell.c };
            }
            this._currentTx = {
                viewInfo,
                actions: [action]
            };

            if (!this._manualMode && !this._microTaskPending) {
                this._microTaskPending = true;
                queueMicrotask(() => this._autoCommit());
            }
        }

        this._redoStack = [];
    }

    /** @param {string} [name=''] @returns {void} Begin a manual undo transaction. */
    begin(name = '') {
        if (this._manualDepth > 0) {
            this._manualDepth += 1;
            return;
        }
        this._manualMode = true;
        this._manualDepth = 1;
        if (this._currentTx) {
            if (name && !this._currentTx.name) this._currentTx.name = name;
            return;
        }
        this._currentTx = {
            name,
            viewInfo: this._captureView(),
            actions: []
        };
    }

    /** @returns {void} Commit current manual transaction. */
    commit() {
        if (this._manualDepth === 0) return;
        this._manualDepth -= 1;
        if (this._manualDepth > 0) return;
        this._manualMode = false;
        this._commitCurrentTx();
    }

    /** @returns {boolean} */
    undo() {
        if (this._currentTx && !this._manualMode) {
            this._commitCurrentTx();
        }

        if (this._undoStack.length === 0) return false;

        const tx = this._undoStack.pop();
        if (this._SN.Event.hasListeners('beforeUndo')) {
            const beforeEvent = this._SN.Event.emit('beforeUndo', { tx });
            if (beforeEvent.canceled) {
                this._undoStack.push(tx);
                return false;
            }
        }
        if (!this._restoreView(tx.viewInfo)) {
            this._undoStack.push(tx);
            return false;
        }

        for (let i = tx.actions.length - 1; i >= 0; i--) {
            tx.actions[i].undo();
        }

        this._redoStack.push(tx);
        this._afterUndoRedo();
        if (this._SN.Event.hasListeners('afterUndo')) {
            this._SN.Event.emit('afterUndo', { tx });
        }
        return true;
    }

    /** @returns {boolean} */
    redo() {
        if (this._redoStack.length === 0) return false;

        const tx = this._redoStack.pop();
        if (this._SN.Event.hasListeners('beforeRedo')) {
            const beforeEvent = this._SN.Event.emit('beforeRedo', { tx });
            if (beforeEvent.canceled) {
                this._redoStack.push(tx);
                return false;
            }
        }
        if (!this._restoreView(tx.viewInfo)) {
            this._redoStack.push(tx);
            return false;
        }

        for (const action of tx.actions) {
            action.redo();
        }

        this._undoStack.push(tx);
        this._afterUndoRedo();
        if (this._SN.Event.hasListeners('afterRedo')) {
            this._SN.Event.emit('afterRedo', { tx });
        }
        return true;
    }

    /** @returns {void} Clear undo/redo stacks and pending transaction. */
    clear() {
        this._undoStack = [];
        this._redoStack = [];
        this._currentTx = null;
        this._manualMode = false;
        this._manualDepth = 0;
        this._microTaskPending = false;
    }

    _autoCommit() {
        this._microTaskPending = false;
        if (!this._manualMode) {
            this._commitCurrentTx();
        }
    }

    _commitCurrentTx() {
        if (!this._currentTx || this._currentTx.actions.length === 0) {
            this._currentTx = null;
            return;
        }

        this._undoStack.push(this._currentTx);
        if (this._undoStack.length > this.maxStack) {
            this._undoStack.shift();
        }
        this._currentTx = null;
    }

    _afterUndoRedo() {
        if (this._SN.DependencyGraph && this._SN.calcMode !== 'manual') {
            this._SN.DependencyGraph.recalculateVolatile();
        }
        this._SN._r();
    }

    _captureView() {
        const sheet = this._SN.activeSheet;
        const vi = sheet.vi;
        return {
            sheetName: sheet.name,
            c: vi?.colArr?.[0] ?? 0,
            r: vi?.rowArr?.[0] ?? 0,
            activeCell: { r: sheet.activeCell.r, c: sheet.activeCell.c }
        };
    }

    _restoreView(viewInfo) {
        if (!viewInfo?.sheetName) return true;

        const targetSheet = this._SN.getSheet(viewInfo.sheetName);
        if (!targetSheet) return false;

        if (this._SN.activeSheet.name !== targetSheet.name) {
            this._SN.activeSheet = targetSheet;
        }
        const sheet = this._SN.activeSheet;
        sheet.viewStart = { r: viewInfo.r, c: viewInfo.c };
        sheet.activeCell = { r: viewInfo.activeCell.r, c: viewInfo.activeCell.c };
        sheet._activeAreas = [];

        return true;
    }
}
