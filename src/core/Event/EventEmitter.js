/**
 * SheetNext event system
 * High-performance, minimal event transmitter
 */

import { EventPriority } from '../../enum/index.js';
export { EventPriority };

/**
 * Event object
 */
export class SheetEvent {
    /** @param {string} type @param {Object} data @param {Object} target */
    constructor(type, data, target) {
        /** @type {string} */
        this.type = type;
        /** @type {Object} */
        this.data = data;
        /** @type {Object} */
        this.target = target;
        this.canceled = false;
        this.stopped = false;
        this.cancelReason = null;
        /** @type {number} */
        this.timestamp = Date.now();
        this._waitUntil = null;
    }

    /**
     * Cancel event (block default behavior)
     * @param {string} reason - Reason for canceling (optional, used as a reminder)
     */
    cancel(reason = '') {
        this.canceled = true;
        this.cancelReason = reason;
        if (reason && this.target?.Utils?.toast) {
            this.target.Utils.toast(reason);
        }
    }

    /**
     * Stop event propagation (no longer performing follow-up listener)
     */
    stopPropagation() {
        this.stopped = true;
    }

    /**
     * Register async work for emitAsync to await.
     * @param {Promise} promise
     */
    waitUntil(promise) {
        if (!promise || typeof promise.then !== 'function') return;
        if (!this._waitUntil) this._waitUntil = [];
        this._waitUntil.push(promise);
    }
}

/**
 * Event Transmitter
 */
export default class EventEmitter {
    constructor() {
        this._events = new Map();       // { eventName: listeners[] }
        this._onceSet = new WeakSet();  // 标记一次性监听器
    }

    /**
     * Register Event Listener
     * @param {string} event - Event Name
     * @param {Function} handler - Handling function
     * @param {Object} options - Options
     * @param {string} options.id - Unique identity (anti duplicate binding)
     * @param {number} options.priority - Priority (execute as soon as possible)
     * @param {boolean} options.once - Whether to execute only once
     * @returns {EventEmitter} - Chain Calls Supported
     */
    on(event, handler, options = {}) {
        if (typeof handler !== 'function') return this;

        const { id = null, priority = 0, once = false } = options;

        let listeners = this._events.get(event);
        if (!listeners) {
            listeners = [];
            this._events.set(event, listeners);
        }

        // 如果有 id，先移除同 id 的旧监听器（防重复）
        if (id) {
            const idx = listeners.findIndex(l => l.id === id);
            if (idx !== -1) listeners.splice(idx, 1);
        }

        const listener = { handler, id, priority };
        if (once) this._onceSet.add(listener);

        // 二分插入，保持按 priority 升序排列
        const insertIdx = this._binarySearchInsert(listeners, priority);
        listeners.splice(insertIdx, 0, listener);

        return this;
    }

    /**
     * Sign up for a one-time event listener
     * @param {string} event - Event Name
     * @param {Function} handler - Handling function
     * @param {Object} options - Options
     * @returns {EventEmitter}
     */
    once(event, handler, options = {}) {
        return this.on(event, handler, { ...options, once: true });
    }

    /**
     * Remove Event Listener
     * @param {string} event - Event Name
     * @param {string|Function} idOrHandler - id or handler reference
     * @returns {EventEmitter}
     */
    off(event, idOrHandler) {
        const listeners = this._events.get(event);
        if (!listeners) return this;

        if (idOrHandler === undefined) {
            // 移除该事件的所有监听器
            this._events.delete(event);
            return this;
        }

        const idx = listeners.findIndex(l =>
            l.id === idOrHandler || l.handler === idOrHandler
        );
        if (idx !== -1) {
            const [removed] = listeners.splice(idx, 1);
            this._onceSet.delete(removed);
        }

        return this;
    }

    /**
     * Remove all event listeners
     * @returns {EventEmitter}
     */
    offAll() {
        this._events.clear();
        return this;
    }

    /**
     * Trigger event
     * @param {string} event - Event Name
     * @param {Object} data - Event Data
     * @returns {SheetEvent} - Event object (canceled status can be checked)
     */
    emit(event, data = {}) {
        const e = new SheetEvent(event, data, this);
        const listeners = this._events.get(event);

        // 无监听器，直接返回（零开销）
        if (!listeners || listeners.length === 0) return e;

        // 复制数组，防止遍历时修改导致问题
        const arr = listeners.slice();
        const toRemove = [];

        for (const listener of arr) {
            if (e.stopped) break;

            try {
                listener.handler.call(this, e);
            } catch (err) {
                console.error(`[SheetNext] Event handler error (${event}):`, err);
            }

            // 标记一次性监听器待移除
            if (this._onceSet.has(listener)) {
                this._onceSet.delete(listener);
                toRemove.push(listener);
            }
        }

        // 批量移除一次性监听器
        if (toRemove.length > 0) {
            for (const listener of toRemove) {
                const idx = listeners.indexOf(listener);
                if (idx !== -1) listeners.splice(idx, 1);
            }
        }

        return e;
    }

    /**
     * Trigger event with async hooks support.
     * @param {string} event - event name
     * @param {Object} data - event data
     * @returns {Promise<SheetEvent>}
     */
    async emitAsync(event, data = {}) {
        const e = new SheetEvent(event, data, this);
        const listeners = this._events.get(event);

        if (!listeners || listeners.length === 0) return e;

        const arr = listeners.slice();
        const toRemove = [];

        for (const listener of arr) {
            if (e.stopped) break;

            try {
                const res = listener.handler.call(this, e);
                if (res && typeof res.then === 'function') {
                    await res;
                }
            } catch (err) {
                console.error(`[SheetNext] Event handler error (${event}):`, err);
            }

            if (this._onceSet.has(listener)) {
                this._onceSet.delete(listener);
                toRemove.push(listener);
            }
        }

        if (e._waitUntil && e._waitUntil.length > 0) {
            const settled = await Promise.allSettled(e._waitUntil);
            for (const result of settled) {
                if (result.status === 'rejected') {
                    console.error(`[SheetNext] Event waitUntil error (${event}):`, result.reason);
                }
            }
        }

        if (toRemove.length > 0) {
            for (const listener of toRemove) {
                const idx = listeners.indexOf(listener);
                if (idx !== -1) listeners.splice(idx, 1);
            }
        }

        return e;
    }

    /**
     * Check for listeners
     * @param {string} event - Event Name
     * @returns {boolean}
     */
    hasListeners(event) {
        const listeners = this._events.get(event);
        return listeners && listeners.length > 0;
    }

    /**
     * Get the number of listeners
     * @param {string} event - Event name (optional, return total if not passed)
     * @returns {number}
     */
    listenerCount(event) {
        if (event) {
            const listeners = this._events.get(event);
            return listeners ? listeners.length : 0;
        }
        let count = 0;
        for (const listeners of this._events.values()) {
            count += listeners.length;
        }
        return count;
    }

    /**
     * Get all event names and number of monitors
     * @returns {Array<{event: string, count: number}>}
     */
    listEvents() {
        const result = [];
        for (const [event, listeners] of this._events.entries()) {
            result.push({ event, count: listeners.length });
        }
        return result;
    }

    /**
     * Binary Find Insertion Position (ascending priority)
     * @private
     */
    _binarySearchInsert(arr, priority) {
        let lo = 0, hi = arr.length;
        while (lo < hi) {
            const mid = (lo + hi) >>> 1;
            if (arr[mid].priority <= priority) {
                lo = mid + 1;
            } else {
                hi = mid;
            }
        }
        return lo;
    }
}
