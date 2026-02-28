// 辅助函数：颜色插值
export function interpolateColor(color1, color2, ratio) {
    const r1 = parseInt(color1.substr(1, 2), 16);
    const g1 = parseInt(color1.substr(3, 2), 16);
    const b1 = parseInt(color1.substr(5, 2), 16);

    const r2 = parseInt(color2.substr(1, 2), 16);
    const g2 = parseInt(color2.substr(3, 2), 16);
    const b2 = parseInt(color2.substr(5, 2), 16);

    const r = Math.round(r1 + (r2 - r1) * ratio);
    const g = Math.round(g1 + (g2 - g1) * ratio);
    const b = Math.round(b1 + (b2 - b1) * ratio);

    return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

export function shouldHideFormula(sheet, cell) {
    return !!(sheet?.protection?.enabled && cell?.protection?.hidden && cell?.isFormula);
}

export function getFormulaBarValue(sheet, cell) {
    if (!cell) return '';
    if (shouldHideFormula(sheet, cell)) return '';
    return cell.editVal?.toString() ?? '';
}
