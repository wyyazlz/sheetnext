import { CHART_CATEGORIES } from './ChartConfig.js';
import { getChartSvg } from '../assets/chartSvgs.js';

export function buildChartModalHTML(categories) {
    const tabs = categories.map((cat) => (
        `<span class="sn-chart-tab" data-category="${cat.key}">${cat.text}</span>`
    )).join('');

    const panels = categories.map((cat) => {
        const items = cat.items.map((item) => (
            `<div class="sn-chart-item" data-category="${cat.key}" data-item="${item.key}" title="${item.text}">
                <div class="sn-chart-preview">${getChartSvg(item.key)}</div>
                <div class="sn-chart-item-label">${item.text}</div>
            </div>`
        )).join('');
        return `<div class="sn-chart-panel" data-category="${cat.key}">${items}</div>`;
    }).join('');

    return `
        <div class="sn-chart-modal">
            <div class="sn-chart-tabs">${tabs}</div>
            <div class="sn-chart-panels">${panels}</div>
        </div>
    `;
}

export function initChartModalInteraction(bodyEl, initialCategory, initialItem) {
    const modalEl = bodyEl.querySelector('.sn-chart-modal');
    if (!modalEl) return;

    const tabs = Array.from(modalEl.querySelectorAll('.sn-chart-tab'));
    const panels = Array.from(modalEl.querySelectorAll('.sn-chart-panel'));

    const selectItem = (itemEl) => {
        if (!itemEl) return;
        modalEl.querySelectorAll('.sn-chart-item.selected').forEach((el) => el.classList.remove('selected'));
        itemEl.classList.add('selected');
        modalEl.dataset.selectedCategory = itemEl.dataset.category || '';
        modalEl.dataset.selectedItem = itemEl.dataset.item || '';
    };

    const setActiveCategory = (key) => {
        const targetKey = key || CHART_CATEGORIES[0]?.key;
        tabs.forEach((tab) => tab.classList.toggle('active', tab.dataset.category === targetKey));
        panels.forEach((panel) => panel.classList.toggle('active', panel.dataset.category === targetKey));
        modalEl.dataset.selectedCategory = targetKey;

        const activePanel = modalEl.querySelector(`.sn-chart-panel[data-category="${targetKey}"]`);
        if (!activePanel?.querySelector('.sn-chart-item.selected')) {
            selectItem(activePanel?.querySelector('.sn-chart-item'));
        }
    };

    setActiveCategory(initialCategory?.key);

    if (initialItem) {
        const initialItemEl = modalEl.querySelector(`.sn-chart-item[data-item="${initialItem}"]`);
        if (initialItemEl) {
            setActiveCategory(initialItemEl.dataset.category);
            selectItem(initialItemEl);
        }
    }

    modalEl.addEventListener('click', (event) => {
        const tabEl = event.target.closest('.sn-chart-tab');
        if (tabEl) {
            setActiveCategory(tabEl.dataset.category);
            return;
        }
        const itemEl = event.target.closest('.sn-chart-item');
        if (itemEl) selectItem(itemEl);
    });
}

export function getSelectedChartItem(bodyEl) {
    return bodyEl.querySelector('.sn-chart-modal')?.dataset?.selectedItem || '';
}
