const VALIDATION_TYPES = new Set(['list', 'whole', 'decimal', 'textLength', 'date', 'time', 'custom']);
const NEED_OPERATOR_TYPES = new Set(['whole', 'decimal', 'textLength', 'date', 'time']);
const TWO_FORMULA_OPERATORS = new Set(['between', 'notBetween']);

function _normalizeArea(area) {
    if (!area?.s || !area?.e) return null;
    return {
        s: {
            r: Math.min(area.s.r, area.e.r),
            c: Math.min(area.s.c, area.e.c)
        },
        e: {
            r: Math.max(area.s.r, area.e.r),
            c: Math.max(area.s.c, area.e.c)
        }
    };
}

function _resolveAreas(sheet, range) {
    const fallback = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    const source = range ?? fallback;
    const list = Array.isArray(source) ? source : [source];
    const areas = [];

    list.forEach(item => {
        const parsed = typeof item === 'string' ? sheet.rangeStrToNum(item) : item;
        const normalized = _normalizeArea(parsed);
        if (normalized) areas.push(normalized);
    });
    return areas;
}

function _clone(obj) {
    if (obj == null) return obj;
    return JSON.parse(JSON.stringify(obj));
}

function _normalizeListValues(value) {
    if (Array.isArray(value)) {
        return value.map(item => String(item ?? '').trim()).filter(Boolean);
    }
    if (typeof value === 'string') {
        return value
            .split(/[,\n，;；]/)
            .map(item => item.trim())
            .filter(Boolean);
    }
    return [];
}

function _normalizeRule(rule) {
    if (!rule || typeof rule !== 'object') return null;
    const type = String(rule.type || '').trim();
    if (!VALIDATION_TYPES.has(type)) {
        throw new Error('数据验证类型无效');
    }

    const next = {
        type,
        allowBlank: rule.allowBlank !== false,
        showInputMessage: rule.showInputMessage !== false,
        showErrorMessage: rule.showErrorMessage !== false
    };

    if (rule.promptTitle != null) next.promptTitle = String(rule.promptTitle);
    if (rule.prompt != null) next.prompt = String(rule.prompt);
    if (rule.errorTitle != null) next.errorTitle = String(rule.errorTitle);
    if (rule.error != null) next.error = String(rule.error);
    if (rule.errorStyle != null) next.errorStyle = String(rule.errorStyle);

    if (type === 'list') {
        const values = _normalizeListValues(rule.formula1);
        if (!values.length) throw new Error('列表验证需要至少一个候选值');
        next.formula1 = values;
        next.showDropDown = rule.showDropDown !== false;
        return next;
    }

    if (type === 'custom') {
        const formula = String(rule.formula1 ?? '').trim();
        if (!formula) throw new Error('自定义验证需要输入公式');
        next.formula1 = formula;
        return next;
    }

    if (NEED_OPERATOR_TYPES.has(type)) {
        next.operator = String(rule.operator || 'between');
        if (rule.formula1 == null || String(rule.formula1).trim() === '') {
            throw new Error('数据验证最小值/条件值不能为空');
        }
        next.formula1 = rule.formula1;
        if (TWO_FORMULA_OPERATORS.has(next.operator)) {
            if (rule.formula2 == null || String(rule.formula2).trim() === '') {
                throw new Error('使用区间比较时，请输入第二个条件值');
            }
            next.formula2 = rule.formula2;
        } else if (rule.formula2 != null && String(rule.formula2).trim() !== '') {
            next.formula2 = rule.formula2;
        }
    }

    return next;
}

export function setDataValidation(range, rule) {
    const areas = _resolveAreas(this, range);
    if (!areas.length) return { success: false, canceled: true, cellCount: 0 };
    const normalizedRule = _normalizeRule(rule);
    const beforeEvent = this.SN.Event.emit('beforeSetDataValidation', {
        sheet: this,
        areas,
        rule: _clone(normalizedRule)
    });
    if (beforeEvent.canceled) return { success: false, canceled: true, cellCount: 0 };

    return this.SN._withOperation('setDataValidation', {
        sheet: this,
        areas,
        rule: _clone(normalizedRule)
    }, () => {
        let cellCount = 0;
        this.eachCells(areas, (r, c) => {
            this.getCell(r, c).dataValidation = _clone(normalizedRule);
            cellCount++;
        });

        const summary = { success: true, areas, rule: _clone(normalizedRule), cellCount };
        this.SN._recordChange({
            type: 'dataValidation',
            action: 'set',
            sheet: this,
            areas,
            rule: _clone(normalizedRule),
            cellCount
        });
        this.SN.Event.emit('afterSetDataValidation', { sheet: this, ...summary });
        return summary;
    });
}

export function clearDataValidation(range) {
    const areas = _resolveAreas(this, range);
    if (!areas.length) return { success: false, canceled: true, cellCount: 0 };
    const beforeEvent = this.SN.Event.emit('beforeClearDataValidation', { sheet: this, areas });
    if (beforeEvent.canceled) return { success: false, canceled: true, cellCount: 0 };

    return this.SN._withOperation('clearDataValidation', {
        sheet: this,
        areas
    }, () => {
        let cellCount = 0;
        this.eachCells(areas, (r, c) => {
            const cell = this.getCell(r, c);
            if (!cell.dataValidation) return;
            cell.dataValidation = null;
            cellCount++;
        });

        const summary = { success: true, areas, cellCount };
        this.SN._recordChange({
            type: 'dataValidation',
            action: 'clear',
            sheet: this,
            areas,
            cellCount
        });
        this.SN.Event.emit('afterClearDataValidation', { sheet: this, ...summary });
        return summary;
    });
}
