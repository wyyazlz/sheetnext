// ==================== XML 解析相关 ====================

import { getColor } from '../Xml/helpers.js';
import { CONNECTOR_TYPES } from './constants.js';

const VERTICAL_TEXT_VALUES = new Set(['wordArtVert', 'wordArtVertRtl', 'eaVert', 'vert', 'vert270', 'mongolianVert']);

// 从 XML 解析形状配置
export function parseShapeFromXml(xmlDetail, SN, isConnector = false) {
    if (!xmlDetail) return null;

    // xmlDetail 本身就是形状内容，直接取属性
    const spPr = xmlDetail['xdr:spPr'];
    const style = xmlDetail['xdr:style'];
    const txBody = xmlDetail['xdr:txBody'];

    // 解析形状类型
    const prstGeom = spPr?.['a:prstGeom'];
    const shapeType = prstGeom?._$prst || 'rect';

    // 根据 shapeType 判断是否为连接线
    const isActualConnector = CONNECTOR_TYPES.includes(shapeType);
    const isTextBox = !isActualConnector && xmlDetail['xdr:nvSpPr']?.['xdr:cNvSpPr']?._$txBox === '1';

    // 解析样式
    const shapeStyle = {
        fill: isActualConnector ? 'none' : '#FFFFFF',
        stroke: '#000000',
        strokeWidth: 1,
        textColor: '#000000',
        fontSize: 11,
        textDirection: 'horizontal',
        startArrow: 'none',
        endArrow: 'none'
    };
    if (isTextBox) {
        shapeStyle.textAlign = 'l';
        shapeStyle.strokeWidth = 0.75;
    }

    // 从 spPr 解析填充色
    const solidFill = spPr?.['a:solidFill'];
    if (solidFill) {
        const fillColor = getColor(solidFill['a:srgbClr'] || solidFill['a:schemeClr'] || solidFill['a:sysClr'], SN);
        if (fillColor) shapeStyle.fill = fillColor;
    }

    // 从 style 的 fillRef 解析填充
    if (style?.['a:fillRef']) {
        const fillRef = style['a:fillRef'];
        const fillColor = getColor(fillRef['a:schemeClr'] || fillRef['a:srgbClr'], SN);
        if (fillColor && fillRef._$idx !== '0') shapeStyle.fill = fillColor;
    }

    // 从 spPr 解析边框
    const ln = spPr?.['a:ln'];
    if (ln) {
        shapeStyle.strokeWidth = (parseInt(ln._$w) || 9525) / 12700;
        const lnFill = ln['a:solidFill'];
        if (lnFill) {
            const strokeColor = getColor(lnFill['a:srgbClr'] || lnFill['a:schemeClr'] || lnFill['a:sysClr'], SN);
            if (strokeColor) shapeStyle.stroke = strokeColor;
        }

        // 解析箭头（根据实际shapeType判断）
        if (isActualConnector) {
            const headEnd = ln['a:headEnd'];
            if (headEnd?._$type) shapeStyle.startArrow = headEnd._$type;
            const tailEnd = ln['a:tailEnd'];
            if (tailEnd?._$type) shapeStyle.endArrow = tailEnd._$type;
        }
    }

    // 从 style 的 lnRef 解析边框
    if (style?.['a:lnRef']) {
        const lnRef = style['a:lnRef'];
        const lnColor = getColor(lnRef['a:schemeClr'] || lnRef['a:srgbClr'], SN);
        if (lnColor && lnRef._$idx !== '0') shapeStyle.stroke = lnColor;
    }

    const bodyPr = txBody?.['a:bodyPr'];
    if (isVerticalText(bodyPr?._$vert)) shapeStyle.textDirection = 'vertical';

    // 解析文本
    let shapeText = '';
    const paragraphs = toArray(txBody?.['a:p']);
    let hasParagraphAlign = false;
    if (paragraphs.length) {
        const lines = [];
        let firstRun = null;
        const firstParagraph = paragraphs[0];

        paragraphs.forEach((p) => {
            const runs = toArray(p?.['a:r']);
            if (!firstRun && runs.length) firstRun = runs[0];
            lines.push(runs.filter(r => r?.['a:t'] !== undefined).map(r => r['a:t']).join(''));
            if (p?.['a:pPr']?._$algn) {
                hasParagraphAlign = true;
                shapeStyle.textAlign = p['a:pPr']._$algn;
            }
        });

        shapeText = lines.join('\n');

        // 解析字体大小 (优先从 run 中获取)
        if (firstRun?.['a:rPr']?._$sz) {
            shapeStyle.fontSize = parseInt(firstRun['a:rPr']._$sz) / 100;
        } else if (firstParagraph?.['a:endParaRPr']?._$sz) {
            shapeStyle.fontSize = parseInt(firstParagraph['a:endParaRPr']._$sz) / 100;
        }

        // 解析文本颜色
        let hasRunColor = false;
        if (firstRun?.['a:rPr']) {
            const rPr = firstRun['a:rPr'];
            const textColor = getColor(rPr['a:solidFill']?.['a:srgbClr'] || rPr['a:solidFill']?.['a:schemeClr'], SN);
            if (textColor) {
                shapeStyle.textColor = textColor;
                hasRunColor = true;
            }
        }

        // 从 style 的 fontRef 解析字体颜色
        if (!hasRunColor && style?.['a:fontRef']) {
            const fontRef = style['a:fontRef'];
            const fontColor = getColor(fontRef['a:schemeClr'] || fontRef['a:srgbClr'], SN);
            if (fontColor) shapeStyle.textColor = fontColor;
        }
    }

    // Excel 竖排文本框默认从右上向下书写
    if (isTextBox && shapeStyle.textDirection === 'vertical' && !hasParagraphAlign) {
        shapeStyle.textAlign = 'r';
    }

    return { shapeType, shapeStyle, shapeText, isTextBox };
}

// ==================== 导出 XML 相关 ====================

// 构建 shape 的 XML 结构（作为实例方法调用，this 指向 Drawing 对象）
export function buildShapeXml(index) {
    const drawing = this; // 实例方法，从 this 获取
    const shapeType = drawing.shapeType || 'rect';
    const style = drawing.shapeStyle || {};
    const text = drawing.shapeText || '';
    const isConnector = drawing.isConnector;
    const isTextBox = !isConnector && !!drawing.isTextBox;
    const textDirection = style.textDirection === 'vertical' ? 'vertical' : 'horizontal';

    // 计算 EMU 尺寸（位置由外层 xdr:from/xdr:to 控制）
    const SN = drawing.sheet.SN;
    const width = SN.Utils.pxToEmu(drawing.width);
    const height = SN.Utils.pxToEmu(drawing.height);

    // 线条宽度（pt 转 EMU）
    const lineWidth = Math.round((style.strokeWidth || 1) * 12700);

    // 如果是连接线，构建 cxnSp 节点
    if (isConnector) {
        // 构建线条节点
        const lnNode = {
            "_$w": lineWidth.toString()
        };

        // 添加起点箭头
        if (style.startArrow && style.startArrow !== 'none') {
            lnNode["a:headEnd"] = {
                "_$type": style.startArrow
            };
        }

        // 添加终点箭头
        if (style.endArrow && style.endArrow !== 'none') {
            lnNode["a:tailEnd"] = {
                "_$type": style.endArrow
            };
        }

        return {
            "xdr:nvCxnSpPr": {
                "xdr:cNvPr": {
                    "_$id": (index + 1).toString(),
                    "_$name": `连接符 ${index + 1}`
                },
                "xdr:cNvCxnSpPr": ""
            },
            "xdr:spPr": {
                "a:xfrm": {
                    "a:off": { "_$x": "0", "_$y": "0" },
                    "a:ext": { "_$cx": width.toString(), "_$cy": height.toString() }
                },
                "a:prstGeom": {
                    "a:avLst": "",
                    "_$prst": shapeType
                },
                "a:ln": lnNode
            },
            "xdr:style": {
                "a:lnRef": {
                    "a:schemeClr": {
                        "_$val": "accent1"
                    },
                    "_$idx": "1"
                },
                "a:fillRef": {
                    "a:schemeClr": {
                        "_$val": "accent1"
                    },
                    "_$idx": "0"
                },
                "a:effectRef": {
                    "a:schemeClr": {
                        "_$val": "accent1"
                    },
                    "_$idx": "0"
                },
                "a:fontRef": {
                    "a:schemeClr": {
                        "_$val": "tx1"
                    },
                    "_$idx": "minor"
                }
            },
            "_$macro": ""
        };
    }

    const defaultAlign = textDirection === 'vertical' ? 'r' : 'l';
    const paragraphs = buildTextParagraphs(text, style.textAlign || defaultAlign, style.fontSize || 11);
    const bodyPr = {
        "_$vertOverflow": "clip",
        "_$horzOverflow": "clip",
        "_$wrap": "square",
        "_$rtlCol": "0",
        "_$anchor": "t"
    };
    if (textDirection === 'vertical') bodyPr._$vert = 'wordArtVertRtl';

    const spPrNode = {
        "a:xfrm": {
            "a:off": { "_$x": "0", "_$y": "0" },
            "a:ext": { "_$cx": width.toString(), "_$cy": height.toString() }
        },
        "a:prstGeom": {
            "a:avLst": "",
            "_$prst": shapeType
        }
    };
    if (drawing._rotation) {
        spPrNode["a:xfrm"]["_$rot"] = Math.round(drawing._rotation * 60000).toString();
    }
    if (isTextBox) {
        spPrNode["a:solidFill"] = {
            "a:schemeClr": {
                "_$val": "lt1"
            }
        };
        spPrNode["a:ln"] = {
            "_$w": lineWidth.toString(),
            "_$cmpd": "sng",
            "a:solidFill": {
                "a:schemeClr": {
                    "a:shade": {
                        "_$val": "50000"
                    },
                    "_$val": "lt1"
                }
            }
        };
    }

    const result = {
        "xdr:nvSpPr": {
            "xdr:cNvPr": {
                "a:extLst": {
                    "a:ext": {
                        "a16:creationId": {
                            "_$xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
                            "_$id": `{${generateGuid()}}`
                        },
                        "_$uri": "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"
                    }
                },
                "_$id": (index + 1).toString(),
                "_$name": isTextBox
                    ? `文本框 ${index + 1}`
                    : `${shapeType === 'rect' ? '矩形' : shapeType === 'ellipse' ? '椭圆' : '形状'} ${index + 1}`
            },
            "xdr:cNvSpPr": isTextBox ? { "_$txBox": "1" } : ""
        },
        "xdr:spPr": spPrNode,
        "xdr:style": isTextBox ? {
            "a:lnRef": {
                "a:scrgbClr": {
                    "_$r": "0",
                    "_$g": "0",
                    "_$b": "0"
                },
                "_$idx": "0"
            },
            "a:fillRef": {
                "a:scrgbClr": {
                    "_$r": "0",
                    "_$g": "0",
                    "_$b": "0"
                },
                "_$idx": "0"
            },
            "a:effectRef": {
                "a:scrgbClr": {
                    "_$r": "0",
                    "_$g": "0",
                    "_$b": "0"
                },
                "_$idx": "0"
            },
            "a:fontRef": {
                "a:schemeClr": {
                    "_$val": "dk1"
                },
                "_$idx": "minor"
            }
        } : {
            "a:lnRef": {
                "a:schemeClr": {
                    "a:shade": {
                        "_$val": "15000"
                    },
                    "_$val": "accent1"
                },
                "_$idx": "2"
            },
            "a:fillRef": {
                "a:schemeClr": {
                    "_$val": "accent1"
                },
                "_$idx": "1"
            },
            "a:effectRef": {
                "a:schemeClr": {
                    "_$val": "accent1"
                },
                "_$idx": "0"
            },
            "a:fontRef": {
                "a:schemeClr": {
                    "_$val": "lt1"
                },
                "_$idx": "minor"
            }
        },
        "xdr:txBody": {
            "a:bodyPr": bodyPr,
            "a:lstStyle": "",
            "a:p": paragraphs.length === 1 ? paragraphs[0] : paragraphs
        },
        "_$macro": "",
        "_$textlink": ""
    };

    return result;
}

function toArray(value) {
    if (Array.isArray(value)) return value;
    if (value === undefined || value === null || value === '') return [];
    return [value];
}

function isVerticalText(vert) {
    return VERTICAL_TEXT_VALUES.has(String(vert || ''));
}

function buildTextParagraphs(text, textAlign, fontSize) {
    const size = Math.round(fontSize * 100).toString();
    const lines = String(text || '').split('\n');
    const runLang = {
        "_$lang": "zh-CN",
        "_$altLang": "en-US",
        "_$sz": size
    };
    const paraRPr = {
        "_$lang": "zh-CN",
        "_$altLang": "en-US",
        "_$sz": size
    };

    return lines.map((line) => {
        const paragraph = {
            "a:pPr": {
                "_$algn": textAlign
            }
        };
        if (line !== '') {
            paragraph["a:r"] = {
                "a:rPr": runLang,
                "a:t": line
            };
        }
        // OOXML paragraph child sequence must place endParaRPr after runs.
        paragraph["a:endParaRPr"] = paraRPr;
        return paragraph;
    });
}

// 生成 GUID
function generateGuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
        const r = Math.random() * 16 | 0;
        const v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16).toUpperCase();
    });
}
