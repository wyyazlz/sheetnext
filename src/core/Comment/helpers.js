/**
 * 批注工具函数
 * VML解析和构建（简化版）
 */

/**
 * 生成VML绘图XML（用于批注显示）
 * @param {Array} comments - Comment实例数组
 * @returns {Object} VML XML对象
 */
export function buildVmlXml(comments) {
    if (!comments || comments.length === 0) return null;

    // VML是一种较老的矢量图形格式，这里生成简化版
    // 实际Excel的VML包含更多细节，但为性能考虑使用最小化结构
    const shapes = comments.map((comment, index) => {
        const { cellRef, visible, width, height, marginLeft, marginTop } = comment;

        return {
            "_$id": `_x0000_s${1026 + index}`,
            "_$type": "#_x0000_t202",
            "_$style": `position:absolute;margin-left:${marginLeft}pt;margin-top:${marginTop}pt;width:${width}pt;height:${height}pt;z-index:1;visibility:${visible ? 'visible' : 'hidden'}`,
            "_$fillcolor": "#ffffe1",
            "_$strokecolor": "#000000",
            "v:fill": {
                "_$color2": "#ffffe1"
            },
            "v:shadow": {
                "_$on": "t",
                "_$color": "black",
                "_$obscured": "t"
            },
            "v:path": {
                "_$o:connecttype": "none"
            },
            "v:textbox": {
                "_$style": "mso-direction-alt:auto",
                "div": {
                    "_$style": "text-align:left"
                }
            },
            "x:ClientData": {
                "_$ObjectType": "Note",
                "x:MoveWithCells": "",
                "x:SizeWithCells": "",
                "x:Anchor": `${comment.cellPos.c + 1}, 15, ${comment.cellPos.r}, 10, ${comment.cellPos.c + 3}, 15, ${comment.cellPos.r + 3}, 10`,
                "x:AutoFill": "False",
                "x:Row": comment.cellPos.r,
                "x:Column": comment.cellPos.c
            }
        };
    });

    return {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "xml": {
            "_$xmlns:v": "urn:schemas-microsoft-com:vml",
            "_$xmlns:o": "urn:schemas-microsoft-com:office:office",
            "_$xmlns:x": "urn:schemas-microsoft-com:office:excel",
            "o:shapelayout": {
                "_$v:ext": "edit",
                "o:idmap": {
                    "_$v:ext": "edit",
                    "_$data": "1"
                }
            },
            "v:shapetype": {
                "_$id": "_x0000_t202",
                "_$coordsize": "21600,21600",
                "_$o:spt": "202",
                "_$path": "m,l,21600r21600,l21600,xe",
                "v:stroke": {
                    "_$joinstyle": "miter"
                },
                "v:path": {
                    "_$gradientshapeok": "t",
                    "_$o:connecttype": "rect"
                }
            },
            "v:shape": shapes.length === 1 ? shapes[0] : shapes
        }
    };
}

/**
 * 从cellRef计算批注指示器位置（红色三角）
 * @param {Object} cellInfo - 单元格视图信息 {x, y, w, h}
 * @returns {Object} {x, y, size}
 */
export function getIndicatorPosition(cellInfo) {
    const size = 6; // 三角形大小
    return {
        x: cellInfo.x + cellInfo.w - size - 1,
        y: cellInfo.y + 1,
        size
    };
}
