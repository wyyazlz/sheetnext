import CommentItem from './CommentItem.js';

function emitIfListeners(SN, event, data) {
    if (!SN?.Event?.hasListeners?.(event)) return null;
    return SN.Event.emit(event, data);
}

/**
 * Comment 批注管理器
 * 高性能设计：Map索引（cellRef -> CommentItem）+ O(1)查找
 * 架构参考：完全复刻 Drawing 的设计模式
 */
export default class Comment {
    constructor(sheet) {
        /** @type {Sheet} */
        this.sheet = sheet;
        /** @type {Map<string, CommentItem>} */
        this.map = new Map();
        this._authors = ['未知作者'];
    }

    /**
     * 从 xmlObj 解析批注数据（从 Excel 文件导入）
     * @param {Object} xmlObj - Sheet 的 XML 对象
     */
    parse(xmlObj) {
        // 先查找批注文件，而不是依赖legacyDrawing
        const commentsRId = this._findCommentsRId();
        if (!commentsRId) return;

        // 获取批注内容文件
        const commentsPath = this.sheet.SN.Xml.getWsRelsFileTarget(this.sheet.rId, commentsRId);
        const commentsObj = this.sheet.SN.Xml.obj[commentsPath];

        if (!commentsObj) return;

        // 获取VML绘图文件（批注显示属性，可选）
        let vmlObj = null;
        if (xmlObj?.legacyDrawing) {
            const legacyDrawingRId = xmlObj.legacyDrawing['_$r:id'];
            const vmlPath = this.sheet.SN.Xml.getWsRelsFileTarget(this.sheet.rId, legacyDrawingRId);
            vmlObj = this.sheet.SN.Xml.obj[vmlPath];
        }

        // 解析作者列表
        const authorsArr = this.sheet.SN.Utils.objToArr(commentsObj.comments?.authors?.author);
        if (authorsArr.length > 0) {
            this._authors = authorsArr;
        }

        // 解析批注列表
        const commentList = this.sheet.SN.Utils.objToArr(commentsObj.comments?.commentList?.Comment);

        commentList.forEach((commentXml, index) => {
            const cellRef = commentXml._$ref;
            const authorId = parseInt(commentXml._$authorId ?? 0);
            const rawAuthor = this._authors[authorId] ?? '未知作者';

            // 解析文本内容
            let rawText = '';
            const textObj = commentXml.text;

            if (textObj) {
                if (typeof textObj.t === 'string') {
                    rawText = textObj.t;
                } else if (Array.isArray(textObj.r)) {
                    rawText = textObj.r.map(r => r.t ?? '').join('');
                } else if (textObj.t?.['#text']) {
                    rawText = textObj.t['#text'];
                }
            }

            // 规范化作者和文本（处理线程批注格式）
            const author = this._normalizeAuthor(rawAuthor);
            const text = this._normalizeCommentText(rawText);

            // 解析VML属性（如果存在）
            const vmlProps = this._parseVmlProps(vmlObj, index);

            const comment = new CommentItem({
                cellRef,
                author,
                text,
                ...vmlProps
            }, this.sheet);

            this.map.set(cellRef, comment);
        });
    }

    /**
     * 解析VML属性（批注的显示位置和大小）
     * @private
     */
    _parseVmlProps(vmlObj, index) {
        if (!vmlObj) return {};

        // VML文件的结构比较复杂，这里做简化处理
        // 实际Excel的VML包含详细的位置、大小、可见性等信息
        // 为性能考虑，使用默认值
        return {
            visible: false,
            width: 144,
            height: 72,
            marginLeft: 107,
            marginTop: 7
        };
    }

    /**
     * 查找批注文件的RId
     * @private
     */
    _findCommentsRId() {
        const sheetRel = this.sheet.SN.Xml.relationship.find(r => r._$Id === this.sheet.rId);
        if (!sheetRel) return null;

        const targetPath = sheetRel._$Target;
        const rels = this.sheet.SN.Xml.getRelsByTarget(targetPath);
        if (!rels) return null;

        const commentRel = rels.find(r => r._$Type?.includes('comments'));
        return commentRel?._$Id;
    }

    /**
     * 添加批注
     * @param {Object} config - {cellRef, author, text, visible, ...}
     * @returns {CommentItem}
     */
    add(config = {}) {
        const { cellRef } = config;
        if (!cellRef) return null;

        const SN = this.sheet.SN;
        const payload = { sheet: this.sheet, cellRef, config };
        const beforeEvent = emitIfListeners(SN, 'beforeCommentAdd', payload);
        if (beforeEvent?.canceled) return null;
        // 如果已存在，先删除
        if (this.map.has(cellRef)) {
            this.remove(cellRef);
        }

        const comment = new CommentItem({
            cellRef,
            author: config.author ?? '未知作者',
            text: config.text ?? '',
            visible: config.visible ?? false,
            width: config.width,
            height: config.height,
            marginLeft: config.marginLeft,
            marginTop: config.marginTop
        }, this.sheet);

        // 添加作者到列表（去重）
        if (!this._authors.includes(comment.author)) {
            this._authors.push(comment.author);
        }

        this.map.set(cellRef, comment);
        emitIfListeners(SN, 'afterCommentAdd', { ...payload, comment });
        return comment;
    }

    /**
     * 获取批注
     * @param {string} cellRef - 单元格引用 "A1"
     * @returns {CommentItem|null}
     */
    get(cellRef) {
        return this.map.get(cellRef) ?? null;
    }

    /**
     * 删除批注
     * @param {string} cellRef - 单元格引用 "A1"
     * @returns {boolean}
     */
    remove(cellRef) {
        const comment = this.map.get(cellRef) ?? null;
        const payload = { sheet: this.sheet, cellRef, comment };
        const beforeEvent = emitIfListeners(this.sheet.SN, 'beforeCommentRemove', payload);
        if (beforeEvent?.canceled) return false;
        const removed = this.map.delete(cellRef);
        if (removed) {
            emitIfListeners(this.sheet.SN, 'afterCommentRemove', payload);
        }
        return removed;
    }

    /**
     * 获取所有批注列表
     * @returns {CommentItem[]}
     */
    getAll() {
        return Array.from(this.map.values());
    }

    /**
     * 清除所有批注的位置缓存（视图变化时调用）
     */
    clearCache() {
        this.map.forEach(comment => comment._clearCache());
    }

    /**
     * 导出为XML对象（用于Excel导出）
     * 将线程批注转换为普通批注格式
     * @returns {Object|null}
     */
    toXmlObject() {
        if (this.map.size === 0) return null;

        // 构建普通作者列表（过滤掉线程批注的tc=格式）
        const normalAuthors = [];
        const authorIndexMap = new Map(); // 原作者 -> 新索引

        this._authors.forEach(author => {
            const normalAuthor = this._normalizeAuthor(author);
            if (!normalAuthors.includes(normalAuthor)) {
                normalAuthors.push(normalAuthor);
            }
            authorIndexMap.set(author, normalAuthors.indexOf(normalAuthor));
        });

        const commentList = Array.from(this.map.values()).map((comment, index) => {
            const authorId = authorIndexMap.get(comment.author) ?? 0;
            const uid = comment.uid ?? `{${this._generateUUID()}}`;

            // 提取普通批注文本（去除线程批注的提示信息）
            const normalText = this._normalizeCommentText(comment.text);

            return {
                text: { t: normalText },
                "_$ref": comment.cellRef,
                "_$authorId": String(authorId),
                "_$shapeId": `_x0000_s${1026 + index}`,
                "_$xr:uid": uid
            };
        });

        const result = {
            "?xml": {
                "_$version": "1.0",
                "_$encoding": "UTF-8",
                "_$standalone": "yes"
            },
            comments: {
                authors: {
                    author: normalAuthors.length === 1 ? normalAuthors[0] : normalAuthors
                },
                commentList: {
                    comment: commentList.length === 1 ? commentList[0] : commentList
                },
                "_$xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "_$xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "_$mc:Ignorable": "xr",
                "_$xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
            }
        };

        return result;
    }

    /**
     * 规范化作者名（将线程批注的tc=格式转为普通作者）
     * @private
     */
    _normalizeAuthor(author) {
        // tc={GUID} 格式是线程批注标识，转为空字符串（不显示作者）
        if (author && author.startsWith('tc=')) {
            return '';
        }
        return author || '';
    }

    /**
     * 规范化批注文本（提取线程批注中的实际内容）
     * @private
     */
    _normalizeCommentText(text) {
        if (!text) return '';

        // 线程批注格式：[线程批注]\n\n...提示文字...\n\n注释:\n    实际内容
        const marker = '注释:\n';
        const idx = text.indexOf(marker);
        if (idx !== -1) {
            // 提取"注释:"之后的内容，并去除前导空格
            return text.substring(idx + marker.length).replace(/^[\s]+/, '');
        }

        // 如果没有"注释:"标记，检查是否有[线程批注]开头
        if (text.startsWith('[线程批注]')) {
            // 尝试找最后一个换行后的内容
            const lines = text.split('\n');
            const lastLine = lines[lines.length - 1].trim();
            if (lastLine) return lastLine;
        }

        return text;
    }

    /**
     * 生成UUID
     * @private
     */
    _generateUUID() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
            const r = Math.random() * 16 | 0;
            const v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16).toUpperCase();
        });
    }
}
