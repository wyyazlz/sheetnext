/**
 * 下拉菜单组件 - 事件委托模式
 * 支持懒加载面板，无需重复初始化
 */

// 重置主菜单样式
function resetMenuStyles(menu) {
    if (!menu) return;
    menu.style.left = '';
    menu.style.right = '';
    menu.style.top = '';
    menu.style.bottom = '';
    menu.style.marginTop = '';
    menu.style.marginBottom = '';
}

// 重置子菜单样式
function resetSubmenuStyles(submenu) {
    if (!submenu) return;
    submenu.style.left = '';
    submenu.style.right = '';
    submenu.style.top = '';
    submenu.style.bottom = '';
}

// 调整主菜单位置
function adjustMenuPosition(menu) {
    if (!menu) return;
    const rect = menu.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;

    if (rect.right > viewportWidth - 10) {
        menu.style.left = 'auto';
        menu.style.right = '0';
    }
    if (rect.bottom > viewportHeight - 10) {
        menu.style.top = 'auto';
        menu.style.bottom = '100%';
        menu.style.marginTop = '0';
        menu.style.marginBottom = '5px';
    }
}

// 调整子菜单位置
function adjustSubmenuPosition(submenu) {
    if (!submenu) return;
    const rect = submenu.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;

    if (rect.right > viewportWidth - 10) {
        submenu.style.left = 'auto';
        submenu.style.right = '100%';
    }
    if (rect.bottom > viewportHeight - 10) {
        const overflow = rect.bottom - viewportHeight + 10;
        submenu.style.top = `-${overflow}px`;
    }
}

// 主初始化函数 - 事件委托模式
const initializedContainers = new WeakSet();
const submenuCloseTimers = new WeakMap();

export function initDropdown(containerDom) {
    if (initializedContainers.has(containerDom)) return;
    initializedContainers.add(containerDom);

    // 关闭下拉菜单（带动画）
    const closeDropdown = (dropdown) => {
        if (!dropdown || dropdown.classList.contains('closing')) return;
        dropdown.classList.add('closing');
        setTimeout(() => {
            dropdown.classList.remove('active', 'closing');
        }, 200);
    };

    // 关闭子菜单（带动画）
    const closeSubmenu = (submenu) => {
        if (!submenu || submenu.classList.contains('closing')) return;
        submenu.classList.add('closing');
        const timer = setTimeout(() => {
            submenu.classList.remove('active', 'closing');
            submenuCloseTimers.delete(submenu);
        }, 150);
        submenuCloseTimers.set(submenu, timer);
    };

    // 关闭所有下拉
    const closeAllDropdowns = () => {
        containerDom.querySelectorAll('.sn-dropdown.active').forEach(closeDropdown);
    };

    // 打开下拉菜单
    const openDropdown = (dropdown, menu) => {
        // 关闭其他下拉
        containerDom.querySelectorAll('.sn-dropdown.active').forEach(el => {
            if (el !== dropdown) closeDropdown(el);
        });
        resetMenuStyles(menu);
        dropdown.classList.add('active');
        adjustMenuPosition(menu);
    };

    // ==================== 事件委托：下拉菜单 toggle ====================
    containerDom.addEventListener('mousedown', (e) => {
        // 跳过 combobox（有独立逻辑）
        if (e.target.closest('.sn-combobox')) return;

        const toggle = e.target.closest('.sn-dropdown-toggle');
        if (!toggle || !containerDom.contains(toggle)) return;

        const disabledHost = toggle.closest('.sn-item-disabled');
        if (disabledHost && containerDom.contains(disabledHost)) return;

        e.preventDefault();
        // 不调用 stopPropagation，让 ToolbarBuilder 的懒加载菜单渲染也能执行

        const dropdown = toggle.closest('.sn-dropdown');
        const menu = dropdown?.querySelector('.sn-dropdown-menu');
        if (!dropdown || !menu) return;

        const isActive = dropdown.classList.contains('active');
        const keepOpen = dropdown.dataset.keepOpen === 'true';

        if (!isActive) {
            openDropdown(dropdown, menu);
        } else if (!keepOpen) {
            closeDropdown(dropdown);
        }
    }, true);

    // ==================== 事件委托：菜单项点击 ====================
    containerDom.addEventListener('click', (e) => {
        // 跳过 combobox
        if (e.target.closest('.sn-combobox')) return;

        const menuItem = e.target.closest('.sn-dropdown-menu li');
        if (!menuItem || !containerDom.contains(menuItem)) return;

        const disabledHost = menuItem.closest('.sn-item-disabled');
        if (disabledHost && containerDom.contains(disabledHost)) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }

        // 有子菜单的项不处理点击关闭
        if (menuItem.classList.contains('sn-has-submenu')) return;
        if (menuItem.classList.contains('sn-dropdown-divider')) return;

        // 检查是否在子菜单内
        const isInSubmenu = menuItem.closest('.sn-dropdown-submenu');

        // 点击菜单项后关闭下拉
        e.stopPropagation();
        const dropdown = menuItem.closest('.sn-dropdown');
        if (dropdown && dropdown.dataset.keepOpen !== 'true') {
            closeDropdown(dropdown);
        }
    });

    // ==================== 事件委托：子菜单悬停 ====================
    containerDom.addEventListener('mouseenter', (e) => {
        const hasSubmenu = e.target.closest('.sn-has-submenu');
        if (!hasSubmenu || !containerDom.contains(hasSubmenu)) return;

        const submenu = hasSubmenu.querySelector(':scope > .sn-dropdown-submenu');
        if (!submenu) return;

        // 取消关闭定时器
        if (submenuCloseTimers.has(submenu)) {
            clearTimeout(submenuCloseTimers.get(submenu));
            submenuCloseTimers.delete(submenu);
            submenu.classList.remove('closing');
        }

        // 关闭同级其他子菜单
        const parentMenu = hasSubmenu.parentElement;
        parentMenu?.querySelectorAll(':scope > li > .sn-dropdown-submenu.active').forEach(sibling => {
            if (sibling !== submenu) closeSubmenu(sibling);
        });

        // 激活当前子菜单
        resetSubmenuStyles(submenu);
        submenu.classList.add('active');
        adjustSubmenuPosition(submenu);
    }, true);

    containerDom.addEventListener('mouseleave', (e) => {
        const hasSubmenu = e.target.closest('.sn-has-submenu');
        if (!hasSubmenu || !containerDom.contains(hasSubmenu)) return;

        // 检查是否移动到子菜单内
        const relatedTarget = e.relatedTarget;
        if (relatedTarget && hasSubmenu.contains(relatedTarget)) return;

        const submenu = hasSubmenu.querySelector(':scope > .sn-dropdown-submenu');
        if (submenu) {
            closeSubmenu(submenu);
        }
    }, true);

    // ==================== 点击外部关闭 ====================
    let touchStartY = 0, touchStartX = 0, isTouchMove = false;

    document.addEventListener('touchstart', (e) => {
        if (e.target.closest('.sn-dropdown-menu') && containerDom.contains(e.target)) {
            touchStartY = e.touches[0].clientY;
            touchStartX = e.touches[0].clientX;
            isTouchMove = false;
        }
    }, { passive: true });

    document.addEventListener('touchmove', (e) => {
        if (e.target.closest('.sn-dropdown-menu') && containerDom.contains(e.target)) {
            const deltaY = Math.abs(e.touches[0].clientY - touchStartY);
            const deltaX = Math.abs(e.touches[0].clientX - touchStartX);
            if (deltaY > 10 || deltaX > 10) isTouchMove = true;
        }
    }, { passive: true });

    const handleOutsideClick = (e) => {
        if (e.type === 'touchend' && isTouchMove) {
            isTouchMove = false;
            return;
        }

        // combobox 有独立逻辑
        if (e.target.closest('.sn-combobox') && containerDom.contains(e.target.closest('.sn-combobox'))) return;

        const clickedMenu = e.target.closest('.sn-dropdown-menu');
        const clickedToggle = e.target.closest('.sn-dropdown-toggle');

        if (clickedMenu && containerDom.contains(clickedMenu)) {
            // 点击菜单内部，检查 keep-open
            const dropdown = clickedMenu.closest('.sn-dropdown');
            if (dropdown?.dataset.keepOpen !== 'true') {
                // 由菜单项点击事件处理
            }
        } else if (!clickedToggle || !containerDom.contains(clickedToggle)) {
            // 点击外部，关闭所有
            closeAllDropdowns();
        }
    };

    window.addEventListener('click', handleOutsideClick);
    window.addEventListener('touchend', handleOutsideClick);

    // ==================== 窗口 resize ====================
    window.addEventListener('resize', () => {
        containerDom.querySelectorAll('.sn-dropdown.active .sn-dropdown-menu').forEach(menu => {
            resetMenuStyles(menu);
            adjustMenuPosition(menu);
        });
    });

    // ==================== Combobox 事件委托 ====================
    // 打开 combobox
    const openCombobox = (combobox) => {
        if (combobox.classList.contains('active')) return;
        containerDom.querySelectorAll('.sn-dropdown.active').forEach(el => {
            if (el !== combobox) closeDropdown(el);
        });
        const menu = combobox.querySelector('.sn-dropdown-menu');
        resetMenuStyles(menu);
        combobox.classList.add('active');
        adjustMenuPosition(menu);
    };

    // 重置过滤
    const resetComboboxFilter = (combobox) => {
        const menu = combobox.querySelector('.sn-dropdown-menu');
        if (!menu) return;
        menu.querySelectorAll('.sn-dropdown-divider').forEach(d => d.style.display = '');
        menu.querySelectorAll('li:not(.sn-dropdown-divider)').forEach(item => item.style.display = '');
    };

    // 过滤菜单项
    const filterComboboxItems = (combobox, keyword) => {
        const menu = combobox.querySelector('.sn-dropdown-menu');
        if (!menu) return;
        const lowerKeyword = keyword.toLowerCase().trim();
        const items = menu.querySelectorAll('li:not(.sn-dropdown-divider)');
        const dividers = menu.querySelectorAll('.sn-dropdown-divider');

        if (!lowerKeyword) {
            resetComboboxFilter(combobox);
            return;
        }
        dividers.forEach(d => d.style.display = 'none');
        items.forEach(item => {
            item.style.display = item.textContent.toLowerCase().includes(lowerKeyword) ? '' : 'none';
        });
    };

    // Combobox main 区域点击
    containerDom.addEventListener('mousedown', (e) => {
        const comboboxMain = e.target.closest('.sn-combobox-main');
        if (!comboboxMain) return;

        const combobox = comboboxMain.closest('.sn-combobox');
        if (!combobox || !containerDom.contains(combobox)) return;
        if (combobox.classList.contains('sn-item-disabled')) return;

        const input = combobox.querySelector('.sn-combobox-input');

        if (e.target === input) {
            e.stopPropagation();
            resetComboboxFilter(combobox);
            openCombobox(combobox);
            return;
        }

        e.preventDefault();
        e.stopPropagation();
        resetComboboxFilter(combobox);
        openCombobox(combobox);
        if (input) {
            input.focus();
            input.setSelectionRange(input.value.length, input.value.length);
        }
    });

    // Combobox 箭头点击
    containerDom.addEventListener('mousedown', (e) => {
        const arrow = e.target.closest('.sn-combobox-arrow');
        if (!arrow) return;

        const combobox = arrow.closest('.sn-combobox');
        if (!combobox || !containerDom.contains(combobox)) return;
        if (combobox.classList.contains('sn-item-disabled')) return;

        e.preventDefault();
        e.stopPropagation();
        resetComboboxFilter(combobox);

        if (combobox.classList.contains('active')) {
            closeDropdown(combobox);
        } else {
            openCombobox(combobox);
        }
    });

    // Combobox 输入过滤
    containerDom.addEventListener('input', (e) => {
        const input = e.target.closest('.sn-combobox-input');
        if (!input) return;

        const combobox = input.closest('.sn-combobox');
        if (!combobox || !containerDom.contains(combobox)) return;
        if (combobox.classList.contains('sn-item-disabled')) return;

        openCombobox(combobox);
        filterComboboxItems(combobox, input.value);
    });

    // Combobox 菜单项选择
    containerDom.addEventListener('click', (e) => {
        const combobox = e.target.closest('.sn-combobox');
        if (!combobox || !containerDom.contains(combobox)) return;
        if (combobox.classList.contains('sn-item-disabled')) return;

        const menuItem = e.target.closest('.sn-dropdown-menu li');
        if (!menuItem || menuItem.classList.contains('sn-dropdown-divider')) return;

        e.stopPropagation();
        const input = combobox.querySelector('.sn-combobox-input');
        if (input) {
            const text = menuItem.querySelector('span')?.textContent || menuItem.textContent;
            input.value = text.trim();
        }
        resetComboboxFilter(combobox);
        closeDropdown(combobox);
    });

    // Combobox 键盘操作
    containerDom.addEventListener('keydown', (e) => {
        const input = e.target.closest('.sn-combobox-input');
        if (!input) return;

        const combobox = input.closest('.sn-combobox');
        if (!combobox || !containerDom.contains(combobox)) return;
        if (combobox.classList.contains('sn-item-disabled')) return;

        if (e.key === 'Enter') {
            e.preventDefault();
            const menu = combobox.querySelector('.sn-dropdown-menu');
            const visibleItem = menu?.querySelector('li:not(.sn-dropdown-divider):not([style*="display: none"])');
            if (visibleItem) visibleItem.click();
        } else if (e.key === 'Escape') {
            closeDropdown(combobox);
            input.blur();
        }
    });
}
