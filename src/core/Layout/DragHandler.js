/**
 * 拖拽处理模块
 * 负责处理AI聊天窗口的拖拽功能
 */

import { calculateDragBounds, constrainToBounds } from './helpers.js';

/**
 * 初始化拖拽事件
 */
export function initDragEvents() {
    const chatHead = this.snChat.querySelector('.sn-chat-head');
    if (chatHead) {
        chatHead.addEventListener('mousedown', (e) => {
            if (!this.showAIChatWindow) return;
            const target = e.target?.closest ? e.target : e.target?.parentElement;
            if (!target) return;
            if (target.closest('.sn-svg-btn')) return;

            const startX = e.clientX;
            const startY = e.clientY;

            // 获取容器和对话框的位置信息
            const mainRect = this.snMain.getBoundingClientRect();
            const opRect = this.snOp.getBoundingClientRect();
            const chatRect = this.snChat.getBoundingClientRect();

            // 计算边界
            const bounds = calculateDragBounds(mainRect, opRect, chatRect);

            const move = (ev) => {
                let dx = ev.clientX - startX;
                let dy = ev.clientY - startY;

                let x = this.chatWindowOffsetX + dx;
                let y = this.chatWindowOffsetY + dy;

                // 边界限制：不得超出snMain容器边缘
                const constrained = constrainToBounds(x, y, bounds);
                this.snChat.style.transform = `translate(${constrained.x}px, ${constrained.y}px)`;
            };

            const up = (ev) => {
                this.chatWindowOffsetX += ev.clientX - startX;
                this.chatWindowOffsetY += ev.clientY - startY;

                // 边界限制：不得超出snMain容器边缘
                const constrained = constrainToBounds(this.chatWindowOffsetX, this.chatWindowOffsetY, bounds);
                this.chatWindowOffsetX = constrained.x;
                this.chatWindowOffsetY = constrained.y;

                window.removeEventListener('mousemove', move);
                window.removeEventListener('mouseup', up);
            };

            window.addEventListener('mousemove', move);
            window.addEventListener('mouseup', up);
        });
    }
}
