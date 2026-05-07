import DrawingItem from './DrawingItem.js';

/**
 * Drawing manager - CRUD, XML parse & export
 * @class
 */
export default class Drawing {
    /** @param {Sheet} sheet */
    constructor(sheet) {
        /** @type {Sheet} */
        this.sheet = sheet;
        /** @type {Map<string, DrawingItem>} */
        this.map = new Map();
        this._list = [];
    }

    /** @param {Object} chartOption @param {DrawingOptions} [options] @returns {DrawingItem} */
    addChart(chartOption, options) {
        return this._add({ ...options, type: 'chart', chartOption });
    }

    /** @param {string} imageBase64 @param {DrawingOptions} [options] @returns {DrawingItem} */
    addImage(imageBase64, options) {
        return this._add({ ...options, type: 'image', imageBase64 });
    }

    /** @param {string} shapeType @param {DrawingOptions} [options] @returns {DrawingItem} */
    addShape(shapeType, options) {
        return this._add({ ...options, type: 'shape', shapeType });
    }

    _add(config) {
        // type validators (shape has defaults)
        const validators = {
            chart: () => config.chartOption ? null : '必要参数缺失：chartOption',
            image: () => config.imageBase64 ? null : '必要参数缺失：imageBase64',
            shape: () => null,
            slicer: () => null
        };

        if (!validators[config.type]) {
            return this.sheet.SN.Utils.toast(this.sheet.SN.t('core.drawing.drawing.toast.unsupportedType', { type: config.type }));
        }

        const error = validators[config.type]();
        if (error) {
            const nameMap = { chart: '图表', image: '图片', shape: '形状', slicer: '切片器' };
            const label = nameMap[config.type] || '图形';
            return this.sheet.SN.Utils.toast(this.sheet.SN.t('core.drawing.drawing.toast.createFailedByType', { label, error }));
        }

        if (this.sheet.SN.Event.hasListeners('beforeDrawingAdd')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeDrawingAdd', { sheet: this.sheet, config });
            if (beforeEvent.canceled) return null;
        }

        const drawing = new DrawingItem(config, this.sheet);

        this.sheet._idCount = Math.max(this.sheet._idCount, parseInt(drawing.id.replace('rId', '')) || 0);

        this.map.set(drawing.id, drawing);
        this._list.push(drawing);

        if (this.sheet.SN.Event.hasListeners('afterDrawingAdd')) {
            this.sheet.SN.Event.emit('afterDrawingAdd', { sheet: this.sheet, config, drawing });
        }
        return drawing;
    }

    /** @param {string} id */
    remove(id) {
        const drawing = this.map.get(id);
        if (!drawing) return;
        if (this.sheet.SN.Event.hasListeners('beforeDrawingRemove')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeDrawingRemove', { sheet: this.sheet, drawing, id });
            if (beforeEvent.canceled) return;
        }

        this.map.delete(id);

        const idx = this._list.indexOf(drawing);
        if (idx > -1) this._list.splice(idx, 1);

        if (this.sheet.SN.Event.hasListeners('afterDrawingRemove')) {
            this.sheet.SN.Event.emit('afterDrawingRemove', { sheet: this.sheet, drawing, id });
        }
    }

    /** @param {string} id @returns {DrawingItem|null} */
    get(id) {
        return this.map.get(id) || null;
    }

    /** @param {Object|string} cellRef @returns {DrawingItem[]} */
    getByCell(cellRef) {
        if (typeof cellRef === 'string') {
            cellRef = this.sheet.Utils.cellStrToNum(cellRef);
        }

        return this._list.filter(d => {
            const { s, e } = d.area;
            return cellRef.r >= s.r && cellRef.r <= e.r
                && cellRef.c >= s.c && cellRef.c <= e.c;
        });
    }

    /** @returns {DrawingItem[]} */
    getAll() {
        return this._list;
    }

    // ==================== Internal ====================

    /** @param {Object} xmlObj */
    _parse(xmlObj) {
        if (!xmlObj?.drawing) return;

        const drawingRId = xmlObj.drawing['_$r:id'];
        const dRelsLink = this.sheet.SN.Xml.getWsRelsFileTarget(this.sheet.rId, drawingRId);
        const drawInfo = this.sheet.SN.Xml.obj[dRelsLink];

        if (!drawInfo) return;

        // anchor types: oneCell / twoCell / absolute
        const oneCellAnchors = this.sheet.SN.Utils.objToArr(drawInfo['xdr:wsDr']['xdr:oneCellAnchor']).map(a => ({...a, __anchorType: 'oneCell'}));
        const twoCellAnchors = this.sheet.SN.Utils.objToArr(drawInfo['xdr:wsDr']['xdr:twoCellAnchor']).map(a => {
            const editAs = a['_$editAs'];
            const anchorType = editAs === 'absolute' ? 'absolute' : editAs === 'oneCell' ? 'oneCell' : 'twoCell';
            return {...a, __anchorType: anchorType};
        });
        const absoluteAnchors = this.sheet.SN.Utils.objToArr(drawInfo['xdr:wsDr']['xdr:absoluteAnchor']).map(a => ({...a, __anchorType: 'absolute'}));
        const anchors = [].concat(oneCellAnchors, twoCellAnchors, absoluteAnchors);

        const details = this.sheet.SN.Utils.objToArr(this.sheet.SN.Xml.getRelsByTarget(dRelsLink));

        anchors.forEach(anchor => {
            const instance = this._parseAnchor(anchor, details);
            if (instance) {
                this.map.set(instance.id, instance);
                this._list.push(instance);
            }
        });
    }

    /** @param {Object} anchor @param {Array} details @returns {DrawingItem|null} */
    _parseAnchor(anchor, details) {
        let type, id, detail;
        const rels = this.sheet.SN.Utils.objToArr(details);

        // 1. resolve drawing type
        const resolveGraphicFrame = () => {
            if (anchor['xdr:graphicFrame']) return anchor['xdr:graphicFrame'];
            const alt = anchor['mc:AlternateContent'];
            if (!alt) return null;
            const choices = this.sheet.SN.Utils.objToArr(alt['mc:Choice']);
            for (const choice of choices) {
                if (choice?.['xdr:graphicFrame']) return choice['xdr:graphicFrame'];
            }
            return null;
        };
        const graphicFrame = resolveGraphicFrame();

        if (graphicFrame) {
            const graphicData = graphicFrame['a:graphic']?.['a:graphicData'];
            const uri = String(graphicData?.['_$uri'] || '').toLowerCase();

            // slicer is parsed in Slicer module
            if (uri.includes('/slicer')) return null;

            const chartNode = graphicData?.['c:chart'] || graphicData?.['cx:chart'];
            id = chartNode?.['_$r:id'];
            if (!id) return null;

            type = 'chart';
            const rel = rels.find(d => d._$Id === id);
            const relTarget = rel ? rel._$Target.replace('..', 'xl') : null;
            detail = relTarget ? this.sheet.SN.Xml.obj[relTarget] : null;
        }
        else if (anchor['xdr:pic']) {
            type = 'image';
            id = anchor['xdr:pic']?.['xdr:blipFill']?.['a:blip']?.['_$r:embed'];
            if (!id) return null;
            const rel = rels.find(d => d._$Id === id);
            const relTarget = rel ? rel._$Target.replace('..', 'xl') : null;
            detail = relTarget ? this.sheet.SN.Xml.obj[relTarget] : null;
        }
        else if (anchor['xdr:cxnSp']) {
            type = 'shape'; // connector
            const cxn = anchor['xdr:cxnSp'];
            detail = cxn;
            id = cxn['xdr:nvCxnSpPr']['xdr:cNvPr']['_$id'];
        }
        else if (anchor['xdr:sp']) {
            type = 'shape';
            const sp = anchor['xdr:sp'];
            detail = sp;
            id = sp['xdr:nvSpPr']['xdr:cNvPr']['_$id'];
        }

        if (!type) return null;

        // 2. create DrawingItem
        const drawing = new DrawingItem({
            type,
            startCell: anchor['xdr:from'] ?
                { r: +anchor['xdr:from']['xdr:row'], c: +anchor['xdr:from']['xdr:col'] } :
                { r: 0, c: 0 },
            xmlDetail: detail,
            anchorType: anchor.__anchorType
        }, this.sheet);

        // rotation: Excel stores in 1/60000 deg
        const spPr = anchor['xdr:pic']?.['xdr:spPr'] || anchor['xdr:sp']?.['xdr:spPr'] || anchor['xdr:cxnSp']?.['xdr:spPr'];
        const rot = spPr?.['a:xfrm']?.['_$rot'];
        if (rot) drawing._rotation = +rot / 60000;

        // 3. parse position
        if (anchor.__anchorType === 'absolute' && anchor['xdr:pos']) {
            const posX = this.sheet.Utils.emuToPx(+anchor['xdr:pos']['_$x']);
            const posY = this.sheet.Utils.emuToPx(+anchor['xdr:pos']['_$y']);
            const width = this.sheet.Utils.emuToPx(+anchor['xdr:ext']['_$cx']);
            const height = this.sheet.Utils.emuToPx(+anchor['xdr:ext']['_$cy']);

            drawing._area = {
                s: { r: 0, c: 0, offsetX: posX, offsetY: posY },
                e: { r: 0, c: 0, offsetX: posX + width, offsetY: posY + height }
            };
            drawing._width = width;
            drawing._height = height;
            drawing._offsetX = posX;
            drawing._offsetY = posY;
        } else {
            // oneCell/twoCell: use from/to
            drawing._area = {
                s: {
                    r: +anchor['xdr:from']['xdr:row'],
                    c: +anchor['xdr:from']['xdr:col'],
                    offsetX: this.sheet.Utils.emuToPx(+anchor['xdr:from']['xdr:colOff']),
                    offsetY: this.sheet.Utils.emuToPx(+anchor['xdr:from']['xdr:rowOff'])
                },
                e: {
                    r: +anchor['xdr:to']['xdr:row'],
                    c: +anchor['xdr:to']['xdr:col'],
                    offsetX: this.sheet.Utils.emuToPx(+anchor['xdr:to']['xdr:colOff']),
                    offsetY: this.sheet.Utils.emuToPx(+anchor['xdr:to']['xdr:rowOff'])
                }
            };

            drawing._offsetX = drawing._area.s.offsetX;
            drawing._offsetY = drawing._area.s.offsetY;

            // twoCell: extract end cell ref
            if (drawing.anchorType === 'twoCell') {
                drawing._endCell = { r: drawing._area.e.r, c: drawing._area.e.c };
                drawing._endOffsetX = drawing._area.e.offsetX;
                drawing._endOffsetY = drawing._area.e.offsetY;
            }
        }

        // ensure end cell initialized for export
        this.sheet.getCell(drawing._area.e.r, drawing._area.e.c);

        return drawing;
    }

    _clearAllCache() {
        this._list.forEach(drawing => {
            drawing._drawCache = null;
            drawing._renderOption = null;
        });
    }

    _clearChartRefCaches() {
        for (const d of this._list) {
            if (d._hasChartRefs) {
                d._renderOption = null;
                d._drawCache = null;
            }
        }
    }

    // ===== row/col change handlers =====

    /**
     * @param {number} r - row index
     * @param {number} n - count
     */
    _onRowsInserted(r, n) {
        for (const d of this._list) d._adjustForRowInsert(r, n);
    }

    /**
     * @param {number} r - start row index
     * @param {number} n - count
     */
    _onRowsDeleted(r, n) {
        for (const d of this._list) d._adjustForRowDelete(r, n);
    }

    /**
     * @param {number} c - col index
     * @param {number} n - count
     */
    _onColsInserted(c, n) {
        for (const d of this._list) d._adjustForColInsert(c, n);
    }

    /**
     * @param {number} c - start col index
     * @param {number} n - count
     */
    _onColsDeleted(c, n) {
        for (const d of this._list) d._adjustForColDelete(c, n);
    }

    _toXmlObj() {
        if (this._list.length === 0) return null;

        return {
            drawings: this._list,
            count: this._list.length
        };
    }
}
