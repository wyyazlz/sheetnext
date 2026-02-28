/**
 * StateSync - 工具栏复选框与核心API状态双向绑定同步器
 *
 * 设计思路：
 * 1. 配置中 checked 为 getter 函数，面板渲染时调用获取实时状态
 * 2. API 状态变更时，通过 notify 方法同步已渲染的 checkbox DOM
 * 3. 面板按需加载，未渲染时不做无用功，渲染时自动读取最新状态
 */

export default class StateSync {
    constructor(SN) {
        this.SN = SN;
        // id -> { element, getter } 绑定映射
        this.bindings = new Map();
    }

    /**
     * 注册 checkbox 绑定（面板渲染时由 ToolbarBuilder 调用）
     * @param {string} id - checkbox 的唯一标识（对应配置中的 id）
     * @param {HTMLInputElement} element - checkbox DOM 元素
     * @param {Function} getter - 状态获取函数 (ctx) => boolean
     */
    register(id, element, getter) {
        this.bindings.set(id, { element, getter });
    }

    /**
     * 注销绑定（面板销毁时调用，可选）
     * @param {string} id - checkbox 的唯一标识
     */
    unregister(id) {
        this.bindings.delete(id);
    }

    /**
     * 批量注销（面板销毁时调用）
     * @param {string[]} ids - checkbox 的唯一标识数组
     */
    unregisterAll(ids) {
        ids.forEach(id => this.bindings.delete(id));
    }

    /**
     * API 状态变更时调用，同步已渲染的 checkbox
     * @param {string} id - checkbox 的唯一标识
     * @param {boolean} value - 新状态值
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
     * 获取当前状态（面板渲染时调用）
     * @param {Function|boolean} getter - getter 函数或静态布尔值
     * @returns {boolean} 当前状态
     */
    getState(getter) {
        if (typeof getter === 'function') {
            return !!getter(this.SN);
        }
        return !!getter;
    }

    /**
     * 检查是否已注册
     * @param {string} id - checkbox 的唯一标识
     * @returns {boolean}
     */
    has(id) {
        return this.bindings.has(id);
    }

    /**
     * 清空所有绑定
     */
    clear() {
        this.bindings.clear();
    }
}
