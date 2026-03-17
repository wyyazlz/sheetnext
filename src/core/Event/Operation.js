let opSeq = 0;

export class ChangeSet {
    constructor() {
        /** @type {Array<Object>} */
        this.items = [];
        this._summary = null;
    }

    /** @param {Object} change */
    add(change) {
        if (!change) return;
        this.items.push(change);
        this._summary = null;
    }

    /** @type {number} */
    get size() {
        return this.items.length;
    }

    /** @type {{total:number, counts:Object<string, number>}} */
    get summary() {
        if (this._summary) return this._summary;
        const counts = {};
        for (const item of this.items) {
            const type = item?.type || 'unknown';
            counts[type] = (counts[type] || 0) + 1;
        }
        this._summary = { total: this.items.length, counts };
        return this._summary;
    }
}

export class Operation {
    /** @param {string} type @param {Object} [data={}] @param {{parentId?: number|null, meta?: Object|null, trackChanges?: boolean}} [options={}] */
    constructor(type, data = {}, options = {}) {
        /** @type {number} */
        this.id = ++opSeq;
        /** @type {string} */
        this.type = type;
        /** @type {Object} */
        this.data = data;
        /** @type {number} */
        this.startTime = Date.now();
        this.endTime = null;
        this.duration = 0;
        /** @type {number|null} */
        this.parentId = options.parentId ?? null;
        /** @type {Object|null} */
        this.meta = options.meta ?? null;
        /** @type {ChangeSet|null} */
        this.changes = options.trackChanges ? new ChangeSet() : null;
    }

    /** @param {Object} change */
    addChange(change) {
        if (this.changes) this.changes.add(change);
    }

    end() {
        this.endTime = Date.now();
        this.duration = this.endTime - this.startTime;
        return this;
    }
}
