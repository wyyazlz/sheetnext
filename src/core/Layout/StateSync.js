/**
 * StateSync - Toolbar checkbox bidirectionally binds synchronizer with core API status
 *
 * Design Ideas
 * 1. In the configuration, checked is the getter function, and when the panel is rendered, it is called to get the real-time status.
 * 2. Synchronize rendered checkbox Dom via notify method when API status changes
 * 3. Panels load on demand, unrendered without useless work, and automatically read state of update when rendering
 */

export default class StateSync {
    /** @param {import('../Workbook/Workbook.js').default} SN */
    constructor(SN) {
        /** @type {import('../Workbook/Workbook.js').default} */
        this.SN = SN;
        // id -> { element, getter } 绑定映射
        this.bindings = new Map();
    }

    /**
     * Register checkbox binding (called by ToolbarBuilder at panel rendering)
     * @param {string} id - checkbox unique identifier (id in the configuration)
     * @param {HTMLInputElement} element - Checkbox DOM element
     * @param {Function} getter - Status Retrieving Function (ctx) => boolean
     */
    register(id, element, getter) {
        this.bindings.set(id, { element, getter });
    }

    /**
     * Write-off binding (call on panel destruction, optional)
     * @param {string} id - Checkbox's only identifier
     */
    unregister(id) {
        this.bindings.delete(id);
    }

    /**
     * Batch write-off (call when panel is destroyed)
     * @param {string[]} ids - The only identification array for checkbox
     */
    unregisterAll(ids) {
        ids.forEach(id => this.bindings.delete(id));
    }

    /**
     * Call on API status change, sync replayed checkbox
     * @param {string} id - Checkbox's only identifier
     * @param {boolean} value - New Status Value
     */
    notify(id, value) {
        const binding = this.bindings.get(id);
        if (binding?.element) {
            // 面板已渲染，直接更新 DOM
            binding.element.checked = !!value;
        }
        // 面板未渲染时不需要操作，懒加载时 getter 会读取最新值
    }

    /**
     * Retrieving current status (call on panel rendering)
     * @param {Function|boolean} getter - Geter function or static boolean value
     * @returns {bolean} Current status
     */
    getState(getter) {
        if (typeof getter === 'function') {
            return !!getter(this.SN);
        }
        return !!getter;
    }

    /**
     * Check for registration
     * @param {string} id - Checkbox's only identifier
     * @returns {boolean}
     */
    has(id) {
        return this.bindings.has(id);
    }

    /**
     * Clear all bounds.
     */
    clear() {
        this.bindings.clear();
    }
}
