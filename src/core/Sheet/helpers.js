// 数据验证date转js日期字符串
export const _formatDate = f => {

    if (typeof f != 'string') return undefined

    // 先匹配 Excel 内部 DATE(年,月,日) 格式
    const dateMatch = f.match(/DATE\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (dateMatch) {
        const [, y, m, d] = dateMatch;
        const mm = String(m).padStart(2, '0');
        const dd = String(d).padStart(2, '0');
        return `${y}/${mm}/${dd}`;
    }

    // 精确匹配带双引号的 "YYYY/M/D H:MM:SS" 格式
    const dtMatch = f.match(/^"(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})"$/);
    if (dtMatch) {
        const [, y2, m2, d2, h, i, s] = dtMatch;
        const mm2 = String(m2).padStart(2, '0');
        const dd2 = String(d2).padStart(2, '0');
        const hh = String(h).padStart(2, '0');
        const ii = String(i).padStart(2, '0');
        const ss = String(s).padStart(2, '0');
        return `${y2}/${mm2}/${dd2} ${hh}:${ii}:${ss}`;
    }
    // 其他情况原样返回
    return f;
};


export const _formatTime = f => {
    const [, h, m, s] = f.match(/TIME\((\d+),\s*(\d+),\s*(\d+)\)/);
    const hh = String(h).padStart(2, '0');
    const mm = String(m).padStart(2, '0');
    const ss = String(s).padStart(2, '0');
    // 这里以1970-01-01作为默认日期
    return `${hh}:${mm}:${ss}`;
};

// 返回两个区域组成的最大区域，第一个区域固定，第二个区域不分先后
export function findMaxCompositeArea(area1, area2) {
    // 调整 area2 以保证 s 总是左上角，e 总是右下角
    let area2Start = {
        "c": Math.min(area2.s.c, area2.e.c),
        "r": Math.min(area2.s.r, area2.e.r)
    };
    let area2End = {
        "c": Math.max(area2.s.c, area2.e.c),
        "r": Math.max(area2.s.r, area2.e.r)
    };
    // 判断两个区域是否有重叠
    let overlap = !(area1.e.c < area2Start.c || area1.s.c > area2End.c || area1.e.r < area2Start.r || area1.s.r > area2End.r);

    if (!overlap) {
        // 如果没有重叠，返回 area1
        return area2;
    } else {
        // 如果有重叠，合并区域
        return {
            "s": { // start
                "c": Math.min(area1.s.c, area2Start.c), // 列的最小值
                "r": Math.min(area1.s.r, area2Start.r)  // 行的最小值
            },
            "e": { // end
                "c": Math.max(area1.e.c, area2End.c), // 列的最大值
                "r": Math.max(area1.e.r, area2End.r)  // 行的最大值
            }
        };
    }
}