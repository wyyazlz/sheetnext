/**
 * SheetNext 事件系统
 * 高性能、极简的事件发射器
 */

import { EventPriority } from '../../enum/index.js';
export { EventPriority };

/**
 * 事件对象
 */
export class SheetEvent {
    constructor(type, data, target) {
        this.type = type;
        this.data = data;
        this.target = target;
        this.canceled = false;
        this.stopped = false;
        this.cancelReason = null;
        this.timestamp = Date.now();
        this._waitUntil = null;
    }

    /**
     * 取消事件（阻止默认行为）
     * @param {string} reason - 取消原因（可选，用于提示）
     */
    cancel(reason = '') {
        this.canceled = true;
        this.cancelReason = reason;
        if (reason && this.target?.Utils?.toast) {
            this.target.Utils.toast(reason);
        }
    }

    /**
     * 停止事件传播（不再执行后续监听器）
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
 * 事件发射器
 */
export default class EventEmitter {
    constructor() {
        this._events = new Map();       // { eventName: listeners[] }
        this._onceSet = new WeakSet();  // 标记一次性监听器
    }

    /**
     * 注册事件监听器
     * @param {string} event - 事件名
     * @param {Function} handler - 处理函数
     * @param {Object} options - 选项
     * @param {string} options.id - 唯一标识（防重复绑定）
     * @param {number} options.priority - 优先级（越小越先执行）
     * @param {boolean} options.once - 是否只执行一次
     * @returns {EventEmitter} - 支持链式调用
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
     * 注册一次性事件监听器
     * @param {string} event - 事件名
     * @param {Function} handler - 处理函数
     * @param {Object} options - 选项
     * @returns {EventEmitter}
     */
    once(event, handler, options = {}) {
        return this.on(event, handler, { ...options, once: true });
    }

    /**
     * 移除事件监听器
     * @param {string} event - 事件名
     * @param {string|Function} idOrHandler - id 或 handler 引用
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
     * 移除所有事件监听器
     * @returns {EventEmitter}
     */
    offAll() {
        this._events.clear();
        return this;
    }

    /**
     * 触发事件
     * @param {string} event - 事件名
     * @param {Object} data - 事件数据
     * @returns {SheetEvent} - 事件对象（可检查 canceled 状态）
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
     * 检查是否有监听器
     * @param {string} event - 事件名
     * @returns {boolean}
     */
    hasListeners(event) {
        const listeners = this._events.get(event);
        return listeners && listeners.length > 0;
    }

    /**
     * 获取监听器数量
     * @param {string} event - 事件名（可选，不传则返回总数）
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
     * 获取所有事件名与监听数量
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
     * 二分查找插入位置（按 priority 升序）
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
