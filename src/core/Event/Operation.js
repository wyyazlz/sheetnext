let opSeq = 0;

export class ChangeSet {
    constructor() {
        this.items = [];
        this._summary = null;
    }

    add(change) {
        if (!change) return;
        this.items.push(change);
        this._summary = null;
    }

    get size() {
        return this.items.length;
    }

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
    constructor(type, data = {}, options = {}) {
        this.id = ++opSeq;
        this.type = type;
        this.data = data;
        this.startTime = Date.now();
        this.endTime = null;
        this.duration = 0;
        this.parentId = options.parentId ?? null;
        this.meta = options.meta ?? null;
        this.changes = options.trackChanges ? new ChangeSet() : null;
    }

    addChange(change) {
        if (this.changes) this.changes.add(change);
    }

    end() {
        this.endTime = Date.now();
        this.duration = this.endTime - this.startTime;
        return this;
    }
}
