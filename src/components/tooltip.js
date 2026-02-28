let tooltipElement = null;
let currentTarget = null;
let hideTimeout = null;
const initializedContainers = new WeakSet();

function createTooltipElement() {
    tooltipElement = document.createElement('div');
    tooltipElement.className = 'sn-tooltip-container';
    tooltipElement.innerHTML = '<div class="sn-tooltip-arrow"></div><div class="sn-tooltip-content"></div>';
    tooltipElement.setAttribute('data-placement', 'top');
    // 初始化位置在屏幕外，避免占位导致滚动条
    tooltipElement.style.top = '-9999px';
    tooltipElement.style.left = '-9999px';
    document.body.appendChild(tooltipElement);
}

function show(event) {
    clearTimeout(hideTimeout);

    const target = event.currentTarget;
    const title = target.getAttribute('title') || target.getAttribute('data-title');

    if (!title) return;

    if (target.hasAttribute('title')) {
        target.setAttribute('data-title', title);
        target.removeAttribute('title');
    }

    currentTarget = target;

    const contentElement = tooltipElement.querySelector('.sn-tooltip-content');
    contentElement.textContent = title;

    tooltipElement.classList.add('show');

    position(target);
}

function hide(event) {
    hideTimeout = setTimeout(() => {
        tooltipElement.classList.remove('show');
        currentTarget = null;
    }, 100);
}

function position(target) {
    const targetRect = target.getBoundingClientRect();
    const tooltipRect = tooltipElement.getBoundingClientRect();

    const placement = target.getAttribute('data-placement') || 'top';

    let top, left;

    switch (placement) {
        case 'top':
            top = targetRect.top - tooltipRect.height - 8;
            left = targetRect.left + (targetRect.width - tooltipRect.width) / 2;
            tooltipElement.setAttribute('data-placement', 'top');
            break;
        case 'bottom':
            top = targetRect.bottom + 8;
            left = targetRect.left + (targetRect.width - tooltipRect.width) / 2;
            tooltipElement.setAttribute('data-placement', 'bottom');
            break;
        case 'left':
            top = targetRect.top + (targetRect.height - tooltipRect.height) / 2;
            left = targetRect.left - tooltipRect.width - 8;
            tooltipElement.setAttribute('data-placement', 'left');
            break;
        case 'right':
            top = targetRect.top + (targetRect.height - tooltipRect.height) / 2;
            left = targetRect.right + 8;
            tooltipElement.setAttribute('data-placement', 'right');
            break;
        default:
            top = targetRect.top - tooltipRect.height - 8;
            left = targetRect.left + (targetRect.width - tooltipRect.width) / 2;
            tooltipElement.setAttribute('data-placement', 'top');
    }

    const padding = 5;
    if (left < padding) {
        left = padding;
    } else if (left + tooltipRect.width > window.innerWidth - padding) {
        left = window.innerWidth - tooltipRect.width - padding;
    }

    if (top < padding) {
        top = targetRect.bottom + 8;
        tooltipElement.setAttribute('data-placement', 'bottom');
    }

    tooltipElement.style.top = `${top + window.scrollY}px`;
    tooltipElement.style.left = `${left + window.scrollX}px`;
}

function bindEvents(containerDom) {
    const elements = containerDom.querySelectorAll('.sn-tooltip');
    elements.forEach(element => {
        if (!element.dataset.tooltipBound) {
            element.addEventListener('mouseenter', show);
            element.addEventListener('mouseleave', hide);
            element.dataset.tooltipBound = 'true';
        }
    });
}

function observeDOM(containerDom) {
    const observer = new MutationObserver(() => {
        bindEvents(containerDom);
    });

    if (containerDom) {
        observer.observe(containerDom, {
            childList: true,
            subtree: true
        });
    }
}

export function initTooltip(containerDom) {
    // 检查该容器是否已经初始化过
    if (initializedContainers.has(containerDom)) return;
    initializedContainers.add(containerDom);

    if (!tooltipElement) {
        createTooltipElement();
    }

    // 延迟绑定事件，确保 DOM 已经渲染
    setTimeout(() => {
        bindEvents(containerDom);
        observeDOM(containerDom);
    }, 0);
}
