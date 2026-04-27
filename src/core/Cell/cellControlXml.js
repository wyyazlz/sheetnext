export const CELL_CONTROL_TYPES = {
    CHECKBOX: 'checkbox'
};

export const FEATURE_PROPERTY_BAG_PART = 'xl/featurePropertyBag/featurePropertyBag.xml';
export const FEATURE_PROPERTY_BAG_REL_TYPE = 'http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag';
export const FEATURE_PROPERTY_BAG_CONTENT_TYPE = 'application/vnd.ms-excel.featurepropertybag+xml';
export const FEATURE_PROPERTY_BAG_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';
export const XF_COMPLEMENT_EXT_URI = '{C7286773-470A-42A8-94C5-96B5CB345126}';

function _asArray(value) {
    if (value === undefined || value === null) return [];
    return Array.isArray(value) ? value : [value];
}

function _clone(value) {
    if (value === undefined || value === null) return value;
    return JSON.parse(JSON.stringify(value));
}

function _localName(name) {
    const index = String(name).indexOf(':');
    return index === -1 ? name : String(name).slice(index + 1);
}

function _getChild(node, localName) {
    if (!node || typeof node !== 'object') return undefined;
    if (node[localName] !== undefined) return node[localName];
    const key = Object.keys(node).find(item => _localName(item) === localName);
    return key ? node[key] : undefined;
}

function _getChildren(node, localName) {
    return _asArray(_getChild(node, localName));
}

function _readText(node) {
    if (node === undefined || node === null) return undefined;
    if (typeof node === 'object') return node['#text'];
    return node;
}

function _readInt(node, fallback = 0) {
    const value = Number(_readText(node));
    return Number.isFinite(value) ? value : fallback;
}

function _normalizeCheckboxDefault(value) {
    const num = Number(value);
    if (num === 1 || num === 2) return num;
    return 0;
}

export function normalizeCellControl(control) {
    if (!control || typeof control !== 'object') return null;
    if (control.type !== CELL_CONTROL_TYPES.CHECKBOX) return null;
    const result = { type: CELL_CONTROL_TYPES.CHECKBOX };
    if (control.default !== undefined) result.default = _normalizeCheckboxDefault(control.default);
    return result;
}

export function isCheckboxControl(control) {
    return normalizeCellControl(control)?.type === CELL_CONTROL_TYPES.CHECKBOX;
}

function _cellControlSignature(control) {
    const normalized = normalizeCellControl(control);
    if (!normalized) return null;
    return JSON.stringify({
        type: normalized.type,
        default: _normalizeCheckboxDefault(normalized.default)
    });
}

export function ensureCellControlXfComplement(SN, control) {
    const normalized = normalizeCellControl(control);
    if (!normalized) return null;
    if (!SN.buildStyle.cellControls) {
        SN.buildStyle.cellControls = {
            map: new Map(),
            complements: []
        };
    }

    const signature = _cellControlSignature(normalized);
    if (SN.buildStyle.cellControls.map.has(signature)) {
        return SN.buildStyle.cellControls.map.get(signature);
    }

    const index = SN.buildStyle.cellControls.complements.length;
    SN.buildStyle.cellControls.complements.push({
        type: normalized.type,
        default: _normalizeCheckboxDefault(normalized.default)
    });
    SN.buildStyle.cellControls.map.set(signature, index);
    return index;
}

export function applyCellControlXfExtension(xfObj, SN, control) {
    const complementIndex = ensureCellControlXfComplement(SN, control);
    if (complementIndex === null || complementIndex === undefined) return xfObj;

    xfObj.extLst = xfObj.extLst || {};
    const extNode = {
        '_$uri': XF_COMPLEMENT_EXT_URI,
        '_$xmlns:xfpb': FEATURE_PROPERTY_BAG_NS,
        'xfpb:xfComplement': {
            '_$i': String(complementIndex)
        }
    };

    if (!xfObj.extLst.ext) {
        xfObj.extLst.ext = extNode;
        return xfObj;
    }

    const extArr = _asArray(xfObj.extLst.ext)
        .filter(ext => ext?._$uri !== XF_COMPLEMENT_EXT_URI);
    extArr.push(extNode);
    xfObj.extLst.ext = extArr.length === 1 ? extArr[0] : extArr;
    return xfObj;
}

function _getFeaturePropertyBagPart(xml) {
    const direct = xml.obj[FEATURE_PROPERTY_BAG_PART];
    if (direct) return direct;

    const rels = _asArray(xml.relationship);
    const rel = rels.find(item => item?._$Type === FEATURE_PROPERTY_BAG_REL_TYPE);
    if (!rel?._$Target) return null;

    const target = rel._$Target.replace(/^\//, '');
    const key = target.startsWith('xl/') ? target : `xl/${target}`;
    return xml.obj[key] || null;
}

function _getBagTextByKey(bag, key) {
    const item = _getChildren(bag, 'i').find(node => node?._$k === key);
    return _readText(item);
}

function _getBagIdByKey(bag, key) {
    const item = _getChildren(bag, 'bagId').find(node => node?._$k === key);
    if (!item) return undefined;
    return _readInt(item, undefined);
}

function _getMappedBagIds(bag) {
    const mapping = _getChildren(bag, 'a').find(node => node?._$k === 'MappedFeaturePropertyBags');
    return _getChildren(mapping, 'bagId').map(item => _readInt(item, undefined)).filter(Number.isInteger);
}

function _resolveCheckboxControl(bags, xfComplementBagId) {
    const xfComplement = bags[xfComplementBagId];
    if (!xfComplement || xfComplement._$type !== 'XFComplement') return null;

    const xfControlsId = _getBagIdByKey(xfComplement, 'XFControls');
    const xfControls = bags[xfControlsId];
    if (!xfControls || xfControls._$type !== 'XFControls') return null;

    const cellControlId = _getBagIdByKey(xfControls, 'CellControl');
    const cellControl = bags[cellControlId];
    if (!cellControl || cellControl._$type !== 'Checkbox') return null;

    const defaultValue = _normalizeCheckboxDefault(_getBagTextByKey(cellControl, 'default'));
    return { type: CELL_CONTROL_TYPES.CHECKBOX, default: defaultValue };
}

function _buildCellControlStyleMap(xml) {
    const styleSheet = xml.obj?.['xl/styles.xml']?.styleSheet;
    const xfs = _asArray(styleSheet?.cellXfs?.xf);
    const part = _getFeaturePropertyBagPart(xml);
    const root = _getChild(part, 'FeaturePropertyBags');
    const bags = _getChildren(root, 'bag');
    const mapBag = bags.find(bag => bag?._$type === 'XFComplements');
    const mappedBagIds = _getMappedBagIds(mapBag);
    const styleMap = new Map();

    if (!bags.length || !mappedBagIds.length || !xfs.length) return styleMap;

    xfs.forEach((xf, styleIndex) => {
        const ext = _asArray(xf?.extLst?.ext).find(item => item?._$uri === XF_COMPLEMENT_EXT_URI);
        const complementNode = _getChild(ext, 'xfComplement');
        const complementIndex = _readInt(complementNode?._$i, undefined);
        if (!Number.isInteger(complementIndex)) return;

        const bagId = mappedBagIds[complementIndex];
        const control = _resolveCheckboxControl(bags, bagId);
        if (control) styleMap.set(styleIndex, control);
    });

    return styleMap;
}

export function getCellControlForStyle(xml, styleIndex) {
    if (!xml || styleIndex === undefined || styleIndex === null || styleIndex === '') return null;
    if (!xml._cellControlStyleMap) {
        xml._cellControlStyleMap = _buildCellControlStyleMap(xml);
    }
    const index = Number(styleIndex);
    if (!Number.isInteger(index)) return null;
    const control = xml._cellControlStyleMap.get(index);
    return control ? _clone(control) : null;
}

function _ensureOverride(xml) {
    const overrides = _asArray(xml.override);
    const exists = overrides.some(item => item?._$PartName === `/${FEATURE_PROPERTY_BAG_PART}`);
    if (!exists) {
        overrides.push({
            '_$PartName': `/${FEATURE_PROPERTY_BAG_PART}`,
            '_$ContentType': FEATURE_PROPERTY_BAG_CONTENT_TYPE
        });
    }
    xml.obj['[Content_Types].xml'].Types.Override = overrides;
}

function _ensureWorkbookRelationship(xml) {
    const rels = _asArray(xml.relationship);
    const exists = rels.some(item => item?._$Type === FEATURE_PROPERTY_BAG_REL_TYPE);
    if (!exists) {
        rels.push({
            '_$Id': xml._xGetNextRId(),
            '_$Type': FEATURE_PROPERTY_BAG_REL_TYPE,
            '_$Target': FEATURE_PROPERTY_BAG_PART.replace(/^xl\//, '')
        });
    }
    xml.obj['xl/_rels/workbook.xml.rels'].Relationships.Relationship = rels;
}

export function exportCellControlFeaturePropertyBags(SN, xml) {
    const controls = SN.buildStyle?.cellControls?.complements || [];
    if (!controls.length) return false;

    const bags = [];
    const mappedBagIds = [];

    controls.forEach(control => {
        if (!isCheckboxControl(control)) return;

        const checkboxBagId = bags.length;
        const checkboxBag = { '_$type': 'Checkbox' };
        const defaultValue = _normalizeCheckboxDefault(control.default);
        if (defaultValue !== 0) {
            checkboxBag.i = { '_$k': 'default', '#text': String(defaultValue) };
        }
        bags.push(checkboxBag);

        const xfControlsBagId = bags.length;
        bags.push({
            '_$type': 'XFControls',
            'bagId': {
                '_$k': 'CellControl',
                '#text': String(checkboxBagId)
            }
        });

        const xfComplementBagId = bags.length;
        bags.push({
            '_$type': 'XFComplement',
            'bagId': {
                '_$k': 'XFControls',
                '#text': String(xfControlsBagId)
            }
        });

        mappedBagIds.push(xfComplementBagId);
    });

    if (!mappedBagIds.length) return false;

    bags.push({
        '_$type': 'XFComplements',
        '_$extRef': 'XFComplementsMapperExtRef',
        'a': {
            '_$k': 'MappedFeaturePropertyBags',
            'bagId': mappedBagIds.map(id => ({ '#text': String(id) }))
        }
    });

    xml.obj[FEATURE_PROPERTY_BAG_PART] = {
        '?xml': {
            '_$version': '1.0',
            '_$encoding': 'UTF-8',
            '_$standalone': 'yes'
        },
        'FeaturePropertyBags': {
            '_$xmlns': FEATURE_PROPERTY_BAG_NS,
            '_$count': String(bags.length),
            'bag': bags
        }
    };

    _ensureOverride(xml);
    _ensureWorkbookRelationship(xml);
    return true;
}
