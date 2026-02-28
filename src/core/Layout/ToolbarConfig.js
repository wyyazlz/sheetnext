/**
 * Ribbon工具栏配置
 * 定义各面板的工具组和工具项
 */
import { FormulaCatalog } from '../Formula/FormulaCatalog.js';

function applyMinimalTextDefaults(config) {
    config.forEach(panel => {
        (panel.groups || []).forEach(group => applyMinimalTextToItems(group.items || []));
    });
}

function applyMinimalTextToItems(items) {
    items.forEach(item => {
        if ((item.type === 'stack' || item.type === 'row') && Array.isArray(item.items)) {
            applyMinimalTextToItems(item.items);
        }

        if (item.type === 'checkbox' || item.type === 'labelInput') return;
        if ((!item.text && !item.labelKey) || typeof item.minimalText !== 'undefined') return;

        item.minimalText = true;
    });
}

/**
 * 创建工具栏配置
 * @param {string} ns - 命名空间
 * @returns {Object} 工具栏配置
 */
export function createToolbarConfig(ns) {
    const config = [
        // 文件面板
        {
            key: 'file',
            labelKey: 'layout.auto.labelB7394238',
            groups: [
                {
                    items: [
                        { type: 'large', icon: 'daoru', labelKey: 'layout.file.import', action: `${ns}.containerDom.querySelector('.sn-upload').click()` },
                        {
                            type: 'stack', items: [
                                { icon: 'jurassic_addForm', labelKey: 'layout.file.newWorkbook', action: `${ns}.Action.confirmNewWorkbook()` },
                                { icon: 'url', labelKey: 'layout.file.importFromUrl', action: `${ns}.Action.importFromUrl()` },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'daochu', labelKey: 'layout.file.exportXlsx', action: `${ns}.IO.export()` },
                        {
                            type: 'stack', items: [
                                { icon: 'csv', labelKey: 'layout.file.exportCsv', action: `${ns}.IO.export('CSV')` },
                                { icon: 'json', labelKey: 'layout.file.exportJson', action: `${ns}.IO.export('JSON')`, titleKey: 'layout.file.exportJsonTitle' }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { icon: 'html', labelKey: 'layout.file.exportHtml', action: `${ns}.IO.export('HTML')`, titleKey: 'layout.file.exportHtmlTitle' },
                                { icon: 'tp', labelKey: 'layout.file.exportImage', action: `${ns}.IO.exportAllImage()` }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { icon: 'pdf', labelKey: 'layout.file.exportPdf', action: `${ns}.Action.exportPdf()` }
                            ]
                        }
                    ]
                },
                {
                    items: [
                        {
                            type: 'stack', items: [
                                { icon: 'baocun', labelKey: 'layout.file.browserSave', action: `${ns}.Action.saveToLocal()` },
                                { icon: 'huifu', labelKey: 'layout.file.browserRestore', action: `${ns}.Action.loadFromLocal()` },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'dayinyulan', labelKey: 'layout.file.printPreview', action: `${ns}.Action.printPreview()` },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.file.readOnlyAfterExport', id: 'readOnly', checked: ctx => ctx.readOnly, action: `${ns}.readOnly=this.checked` },
                                { icon: 'shuxing', labelKey: 'layout.file.properties', action: `${ns}.Action.showProperties()` }
                            ]
                        }
                    ]
                }
            ]
        },
        // 开始面板
        {
            key: 'start',
            labelKey: 'layout.auto.label396990f1',
            groups: [
                {
                    items: [
                        {
                            type: 'large',
                            icon: 'geshihua',
                            labelKey: 'layout.auto.labelE3362dac',
                            id: 'formatBrush',
                            action: `${ns}.activeSheet.setBrush()`,
                            dblAction: `${ns}.activeSheet.setBrush(true)`,
                            active: ctx => !!ctx.activeSheet?._brush?.area,
                            titleKey: 'layout.auto.title57c65c9c'
                        },
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label3d33e8d4', menu: 'paste' },
                        {
                            type: 'stack', items: [
                                { icon: 'chexiao', titleKey: 'layout.auto.title1c9e8d5c', action: `${ns}.UndoRedo.undo()` },
                                
                                { icon: 'fuzhi', titleKey: 'layout.auto.titleF9d572aa', action: `${ns}.Action.copy()` }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { icon: 'chongzuo', titleKey: 'layout.auto.titleC3cf1168', action: `${ns}.UndoRedo.redo()` },
                                { icon: 'jianqie', titleKey: 'layout.auto.title46adc7a1', action: `${ns}.Action.cut()` },
                            ]
                        },
                    ]
                },
                {
                    items: [
                        {
                            type: 'row', items: [
                                { type: 'dropdown', labelKey: 'layout.auto.label9d40cd00', width: '123px', id: 'fontName', menu: 'fontName' },
                                { type: 'dropdown', labelKey: 'layout.auto.label1ceb2bd8', width: '35px', id: 'fontSize', menu: 'fontSize' },
                                { icon: 'zengjiazihao', action: `${ns}.Action.changeFontSize(true)`, titleKey: 'layout.auto.titleF2e72f6d' },
                                { icon: 'jszh', action: `${ns}.Action.changeFontSize(false)`, titleKey: 'layout.auto.title1a808f64' },
                                
                            ]
                        },
                        {
                            type: 'row', items: [
                                { icon: 'bold', action: `${ns}.Action.fontInversion('bold')`, titleKey: 'layout.auto.titleE3aeedb4' },
                                { icon: 'italic', action: `${ns}.Action.fontInversion('italic')`, titleKey: 'layout.auto.title27c68cf4' },
                                { icon: 'underline', action: `${ns}.Action.fontInversion('underline','single')`, titleKey: 'layout.auto.title7bd6a6c3' },
                                { icon: 'strikethrough', action: `${ns}.Action.fontInversion('strike')`, titleKey: 'layout.auto.titleFd05e5a8' },
                                { type: 'splitDropdown', icon: 'jurassic_borderAll', titleKey: 'layout.auto.title20d59678', action: `${ns}.Action.applyBorderPreset()`, menu: 'border', menuWidth: '180px' },
                                { type: 'splitColorPicker', icon: 'fontColor', titleKey: 'layout.auto.titleFf0455e0', colorType: 'font' },
                                { type: 'splitColorPicker', icon: 'styleFill', titleKey: 'layout.auto.titleB46e89c8', colorType: 'fill' },
                                { type: 'dropdown', icon: 'qingchugeshi1', action: `${ns}.Action.clearAreaFormat()`, menu: 'clear', titleKey: 'layout.auto.titleAb4a78ac' },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        {
                            type: 'row', items: [
                                { icon: 'alignTop', titleKey: 'layout.auto.title2d6a5a8c', action: `${ns}.Action.areasStyle('alignment',{vertical:'top'})` },
                                { icon: 'alignVertically', titleKey: 'layout.auto.titleF2b067c8', action: `${ns}.Action.areasStyle('alignment',{vertical:'center'})` },
                                { icon: 'alignBottom', titleKey: 'layout.auto.title09162f60', action: `${ns}.Action.areasStyle('alignment',{vertical:'bottom'})` },
                                { icon: 'textWrap', titleKey: 'layout.auto.title4fcc8180', action: `${ns}.Action.toggleWrapText()` },
                                { type: 'dropdown', icon: 'ic_zitifangxiang', titleKey: 'layout.auto.titleFaa5e594', menu: 'textRotation' },
                            ]
                        },
                        {
                            type: 'row', items: [
                                { icon: 'alignLeft', titleKey: 'layout.auto.title2f9e586c', action: `${ns}.Action.areasStyle('alignment',{horizontal:'left'})` },
                                { icon: 'alignCenter', titleKey: 'layout.auto.title2724f9ac', action: `${ns}.Action.areasStyle('alignment',{horizontal:'center'})` },
                                { icon: 'alignRight', titleKey: 'layout.auto.title09f772c4', action: `${ns}.Action.areasStyle('alignment',{horizontal:'right'})` },
                                { icon: 'indentIncrease', titleKey: 'layout.auto.title99310054', action: `${ns}.Action.changeIndent(1)` },
                                { icon: 'indentDecrease', titleKey: 'layout.auto.title1cabceb8', action: `${ns}.Action.changeIndent(-1)` },
                            ]
                        },
                        { type: 'large', icon: 'hebing', labelKey: 'layout.auto.labelC602cd4a', hasArrow: true, menu: 'merge' },
                    ]
                },
                {
                    items: [
                        {
                            type: 'row', items: [
                                { type: 'dropdown', labelKey: 'layout.auto.label046ee664', width: '60px', menu: 'numFmt',menuWidth: '200px' },
                                { type: 'dropdown', icon: 'zhuanhuan', labelKey: 'layout.auto.labelF328662a', menu: 'convert' },
                            ]
                        },
                        {
                            type: 'row', items: [
                                { type: 'splitDropdown', icon: 'licai', titleKey: 'layout.auto.title05adf36c', action: `${ns}.Action.areasNumFmt('¥#,##0.00')`, menu: 'currency' },
                                { icon: '_15xiankuantubiaoheji_115', titleKey: 'layout.auto.title047aaa30', action: `${ns}.Action.areasNumFmt('0.00%')` },
                                { icon: 'yunfengdie_tongjishuzhiqianfenwei', titleKey: 'layout.auto.titleD8b83d9c', action: `${ns}.Action.areasNumFmt('#,##0.00')` },
                                { icon: '_16Gongjutubiao2Jianshaoxiaoshuweishu', titleKey: 'layout.auto.title67656060', action: `${ns}.Action.decimalPlaces('cut')` },
                                { icon: '_16Gongjutubiao2Zengjiaxiaoshuweishu', titleKey: 'layout.auto.titleA8781ce4', action: `${ns}.Action.decimalPlaces('add')` },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: 'hanglie', labelKey: 'layout.auto.labelEfe10954', menu: 'rowCol' },
                                { type: 'dropdown', icon: 'quanxianpeizhi', labelKey: 'layout.auto.label1a7e1503', id: 'sheetProtection', menu: 'protection', active: ctx => ctx.activeSheet?.protection?.enabled },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', minimalText: false, icon: 'tiaojiangeshi', labelKey: 'layout.auto.labelB2fe241c', hasArrow: true, menu: 'condFormat',menuWidth: '160px' },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', minimalText: false, icon: 'biaogeyangshi', labelKey: 'layout.auto.label727a825c', menu: 'formatAsTable', menuWidth: '340px' },
                                { type: 'dropdown', minimalText: false, icon: 'danyuangeyangshi', labelKey: 'layout.auto.labelD006b2ec', menu: 'cellStyles', menuWidth: '420px' },
                            ]
                        },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'paixu', labelKey: 'layout.auto.label328b35a0', menu: 'sort' },
                        { type: 'large', icon: 'shaixuan', labelKey: 'layout.auto.labelF3e084a4', menu: 'filter' },
                        { type: 'large', icon: 'dongjie', labelKey: 'layout.auto.label2bb8d6d0', menu: 'freeze' },
                        { type: 'large', icon: 'qiuhe', labelKey: 'layout.auto.label304be598', hasArrow: true, menu: 'autoSum' },
                        { type: 'large', icon: 'chazhao', labelKey: 'layout.auto.label36263538', menu: 'find' }
                    ]
                }
            ]
        },
        // 插入面板
        {
            key: 'insert',
            labelKey: 'layout.auto.label47acfc84',
            groups: [
                {
                    items: [
                        { type: 'large', icon: 'shengchengshujutoushibiao', labelKey: 'layout.auto.labelA05734b8', action: `${ns}.Action.createPivotTable()` },
                        { type: 'large', icon: 'biaoge', labelKey: 'layout.auto.label3dbdf26c', action: `${ns}.Action.createTable()` }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'tupian', labelKey: 'layout.auto.labelCba453e4', action: `${ns}.Action.insertImages()` },
                        { type: 'large', icon: 'jietu', labelKey: 'layout.auto.label2d0ee700', action: `${ns}.Action.insertAreaScreenshot()` },
                        { type: 'large', icon: 'xingzhuang', labelKey: 'layout.auto.label6561bb16', hasArrow: true, menu: 'shapes', menuWidth: '310px' },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'component', labelKey: 'layout.auto.labelC6887918', disabled:true },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'tubiao', labelKey: 'toolbar.insert.chart', action: `${ns}.Action.openChartModal()` },
                        {
                            type: 'row', items: [
                                { type: 'dropdown', icon: 'tubiaoZhuzhuangtu', titleKey: 'chart.category.bar', menu: 'chartBar' },
                                { type: 'dropdown', icon: 'zhexiantu', titleKey: 'chart.category.line', menu: 'chartLine' },
                                { type: 'dropdown', icon: 'bingtu', titleKey: 'chart.category.pie', menu: 'chartPie' },
                            ]
                        },
                        {
                            type: 'row', items: [
                                { type: 'dropdown', icon: 'mianjitu', titleKey: 'chart.category.area', menu: 'chartArea' },
                                { type: 'dropdown', icon: 'sandiantu', titleKey: 'chart.category.scatter', menu: 'chartScatter' },
                                { type: 'dropdown', icon: 'leidatu', titleKey: 'chart.category.radar', menu: 'chartRadar' },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'minitu', labelKey: 'layout.auto.label92d86c42', menu: 'sparkline' },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'chaolianjie', labelKey: 'layout.auto.label72788608', action: `${ns}.Action.insertHyperlink()` },
                        { type: 'large', icon: 'charugongshi', labelKey: 'layout.auto.label92ff62e4', action: `${ns}.Action.openFormulaDialog('all')` },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: '_24_nor', labelKey: 'layout.auto.labelD6d89414', action: `${ns}.Action.insertComment()` },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'wenbenkuang', labelKey: 'layout.auto.labelEd8aca00', hasArrow: true, menu: 'textBox' },
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'huatifuhao', labelKey: 'layout.auto.label9026bf7c',disabled:true },
                    ]
                }
            ]
        },
        // 公式面板
        {
            key: 'formula',
            labelKey: 'layout.auto.label5613e7b9',
            groups: [
                {
                    items: [
                        { type: 'large', icon: 'charugongshi', labelKey: 'layout.auto.label92ff62e4', action: `${ns}.Action.openFormulaDialog('all')` },
                        { type: 'large', icon: 'qiuhe', labelKey: 'layout.auto.label304be598', hasArrow: true, menu: 'autoSum' },
                        { type: 'large', icon: 'icon_fn_fx', labelKey: 'layout.auto.labelD9d58d20', menu: 'commonFn', menuWidth: '220px' },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: 'caiwu', labelKey: 'layout.auto.label20a672e0', menu: 'financeFn' },
                                { type: 'dropdown', icon: 'luojiti', labelKey: 'layout.auto.label92f61dec', menu: 'logicFn' },
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: '_24glFileText', labelKey: 'layout.auto.label5d77ac8c', menu: 'textFn' },
                                { type: 'dropdown', icon: 'calendar', labelKey: 'layout.auto.label92fd0eec', menu: 'dateFn' },
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: 'chaxun', labelKey: 'layout.auto.label36263538', menu: 'lookupFn' },
                                { type: 'dropdown', icon: 'shuxue', labelKey: 'layout.auto.label3a60f922', menu: 'mathFn' },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'vipZidingyimingcheng', labelKey: 'layout.auto.label62002448', action: `${ns}.Action.nameManager()` },
                        {
                            type: 'stack', items: [
                                { icon: 'picimingcheng', labelKey: 'layout.auto.label4b5562f8', action: `${ns}.Action.defineName()` },
                                { icon: 'gongshi', labelKey: 'layout.auto.labelB0c44120', action: `${ns}.Action.useInFormula()` },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        {
                            type: 'stack', items: [
                                { icon: 'zhuizong', labelKey: 'layout.auto.labelF189a560', action: `${ns}.Action.tracePrecedents()` },
                                { icon: 'zhuizongqibeifen', labelKey: 'layout.auto.label071821ca', action: `${ns}.Action.traceDependents()` },
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: 'yichu', labelKey: 'layout.auto.label2d72febc', menu: 'removeArrows' },
                                { icon: 'xianshi', labelKey: 'layout.auto.labelDfebdd0c', action: `${ns}.Action.showFormulas()` },
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { icon: 'jiancha', labelKey: 'layout.auto.label04218143', action: `${ns}.Action.errorCheck()` },
                                { icon: 'jisuan', labelKey: 'layout.auto.labelF3321014', action: `${ns}.Action.evaluateFormula()` },
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'zhihangzhongsuan', labelKey: 'layout.auto.labelFf7d7c40', action: `${ns}.Action.calculateNow()` },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: 'xuanxiang', labelKey: 'layout.auto.label6d1763f6', menu: 'calcOptions' },
                                { icon: 'jisuan', labelKey: 'layout.auto.label9d54bf30', action: `${ns}.Action.calculateSheet()` },
                            ]
                        }
                    ]
                }
            ]
        },
        // 数据面板
        {
            key: 'data',
            labelKey: 'layout.auto.label426d9dba',
            groups: [
                {
                    items: [
                        { type: 'large', icon: 'shengchengshujutoushibiao', labelKey: 'layout.auto.labelA05734b8', action: `${ns}.Action.createPivotTable()` }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'shaixuan', labelKey: 'layout.auto.labelF3e084a4', hasArrow: true, menu: 'filter' },
                        { type: 'large', icon: 'paixu', labelKey: 'layout.auto.label328b35a0', hasArrow: true, menu: 'sort' },
                        {
                            type: 'stack', items: [
                                { icon: 'xianshi', labelKey: 'layout.auto.labelCfc23d16', action: `${ns}.Action.clearFilter()` },
                                { icon: 'gengduo', labelKey: 'layout.auto.label53f4d744', action: `${ns}.Action.reapplyFilter()` }
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'fenlie', labelKey: 'layout.auto.label37d4df3d', action: `${ns}.Action.textToColumns()` },
                        { type: 'large', icon: 'shujuyanzheng', labelKey: 'layout.auto.label065032d4', hasArrow: true, menu: 'dataValidation' },
                        { type: 'large', icon: 'rongqi_2024_10_15T145712642', labelKey: 'layout.auto.label24db7740', action: `${ns}.Action.removeDuplicates()` },
                        { type: 'large', icon: 'shengcheng', labelKey: 'layout.auto.labelEdc074b8', action: `${ns}.Action.flashFill()` }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'chuangjianzu', labelKey: 'layout.auto.labelD27bcc48', action: `${ns}.Action.groupRows()` },
                        { type: 'large', icon: 'ungroup', labelKey: 'layout.auto.label3661dfec', action: `${ns}.Action.ungroupRows()` },
                        { type: 'large', icon: 'fenleihuizong', labelKey: 'layout.auto.label6fbb5db4', action: `${ns}.Action.subtotal()` }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'jisuan', labelKey: 'layout.auto.label32e3aa24', action: `${ns}.Action.consolidate()` },
                        { type: 'large', icon: 'monifenxi', labelKey: 'layout.auto.labelF5b9a3ae', hasArrow: true, menu: 'whatIfAnalysis', disabled: true }
                    ]
                }
            ]
        },
        // 视图面板
        {
            key: 'view',
            labelKey: 'layout.auto.label7d9a5030',
            groups: [
                {
                    items: [
                        { type: 'large', icon: 'huyan_', labelKey: 'layout.auto.label15a1b75c', action: `${ns}.Canvas.toggleEyeProtection()` },
                        { type: 'large', icon: 'dangqianhangliegaoliang', labelKey: 'layout.auto.labelF5e868d8', hasArrow: true, menu: 'highlight' }
                    ]
                },
                {
                    items: [
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label36cb3b80', id: 'showFormulaBar', checked: ctx => ctx.Layout.showFormulaBar, action: `${ns}.Layout.showFormulaBar=this.checked` },
                                { type: 'checkbox', labelKey: 'layout.auto.label805a4b6c', id: 'showGridlines', checked: ctx => ctx.activeSheet?.showGridLines !== false, action: `${ns}.activeSheet.showGridLines=this.checked;${ns}._r()` },
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label5728db81', id: 'showSheetTabBar', checked: ctx => ctx.Layout.showSheetTabBar, action: `${ns}.Layout.showSheetTabBar=this.checked` },
                                { type: 'checkbox', labelKey: 'layout.auto.labelFcc62550', id: 'showAIChat', checked: ctx => ctx.Layout.showAIChat, action: `${ns}.Layout.showAIChat=this.checked` },
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label880beb90', id: 'showStats', checked: ctx => ctx.Layout.showStats, action: `${ns}.Layout.showStats=this.checked` },
                                { type: 'checkbox', labelKey: 'layout.auto.label82b852ee', id: 'fullScreen', checked: () => !!document.fullscreenElement, action: `${ns}.Action.fullScreen()` },
                                
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.labelE8420025', id: 'showRowColHeaders', checked: ctx => ctx.activeSheet?.showRowColHeaders !== false, action: `${ns}.activeSheet.showRowColHeaders=this.checked;${ns}._r()` },
                                {
                                    type: 'checkbox',
                                    labelKey: 'layout.auto.labelCbc4bc74',
                                    id: 'minimalToolbar',
                                    checked: ctx => ctx.Layout.toolbarMode === 'minimal',
                                    disabled: ctx => ctx.Layout.isToolbarModeLocked,
                                    action: `${ns}.Layout.toggleMinimalToolbar(this.checked)`
                                }
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'xianshibili', labelKey: 'layout.auto.label89be961f', hasArrow: true, menu: 'zoom' },
                        { type: 'large', icon: 'yby', labelKey: 'layout.auto.label2f97df3c', action: `${ns}.activeSheet.zoom=1` },
                        { type: 'large', icon: 'suofang', labelKey: 'layout.auto.label8b57557b', action: `${ns}.activeSheet.zoomToSelection()` }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'hanglieqiehuan', labelKey: 'layout.auto.label905e4698',menu:'rowColSize' },
                        {
                            type: 'stack', items: [
                                { icon: 'Edit_dongjiehang_freezeLine_linear', labelKey: 'layout.auto.label12163f5a', action: `${ns}.Action.freezeRows(1)` },
                                { icon: 'Edit_dongjielie_freezeColumn_linear', labelKey: 'layout.auto.labelB99d153f', action: `${ns}.Action.freezeCols(1)` }
                            ]
                        }
                    ]
                }
            ]
        },
        // 页面面板
        {
            key: 'page',
            labelKey: 'layout.auto.labelDbb08c76',
            groups: [
                {
                    items: [
                        { type: 'large', icon: 'zhuti_04', labelKey: 'layout.page.theme', menu: 'theme', disabled: true },
                        {
                            type: 'stack', items: [
                                { type: 'dropdown', icon: 'yanse', labelKey: 'layout.page.colors', menu: 'themeColors', disabled: true },
                                { type: 'dropdown', icon: 'ziti', labelKey: 'layout.page.fonts', menu: 'themeFonts', disabled: true }
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'dayinyulan', labelKey: 'layout.page.printPreview', action: `${ns}.Action.printPreview()` },
                        { type: 'large', icon: 'yebianju', labelKey: 'layout.page.margins', hasArrow: true, menu: 'pageMargins' },
                        { type: 'large', icon: 'fangxiang', labelKey: 'layout.page.orientation', hasArrow: true, menu: 'pageOrientation' },
                        { type: 'large', icon: '_24glFileText', labelKey: 'layout.page.paperSize', hasArrow: true, menu: 'paperSize' },
                        { type: 'large', icon: 'dayinquyuxianxing', labelKey: 'layout.page.printArea', hasArrow: true, menu: 'printArea' },
                        { type: 'large', icon: 'dayinsuofang', labelKey: 'layout.page.printScale', hasArrow: true, menu: 'printScale' },
                        { type: 'large', icon: 'beijingtupian', labelKey: 'layout.page.background', action: `${ns}.Action.setPageBackground()` },

                        {
                            type: 'stack', items: [
                                { icon: 'dayinbiaotihuobiaotou', labelKey: 'layout.page.printTitles', action: `${ns}.Action.setPrintTitles()` },
                                { icon: 'yemeiyejiao_01', labelKey: 'layout.page.headerFooter', action: `${ns}.Action.setHeaderFooter()` }
                            ]
                        }
                    ]
                },
                {
                    items: [
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.page.printGridlines', id: 'printGridlines', checked: ctx => ctx.activeSheet?.printSettings?.printOptions?.gridlines === true, action: `${ns}.Action.togglePrintGridlines(this.checked)` },
                                { type: 'checkbox', labelKey: 'layout.page.printHeadings', id: 'printHeadings', checked: ctx => ctx.activeSheet?.printSettings?.printOptions?.headings === true, action: `${ns}.Action.togglePrintHeadings(this.checked)` }
                            ]
                        }
                    ]
                },
                {
                    items: [
                        { type: 'large', icon: 'dayinyulan1', labelKey: 'layout.page.pagePreview', action: `${ns}.Action.pageBreakPreview()` },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.page.pageBreaks', id: 'showPageBreaks', checked: ctx => ctx.activeSheet?.showPageBreaks === true, action: `${ns}.Action.toggleShowPageBreaks(this.checked)` },
                                { icon: '_24glPageBreak', labelKey: 'layout.page.addBreak', action: `${ns}.Action.insertPageBreak()` }
                            ]
                        }
                    ]
                }
            ]
        },
        // 表设计面板（上下文选项卡，选中表格时显示）
        {
            key: 'tableDesign',
            labelKey: 'layout.auto.labelD4fa30c8',
            contextual: true, // 标记为上下文选项卡
            contextType: 'table', // 上下文类型
            groups: [
                {
                    labelKey: 'layout.auto.label54132264',
                    items: [
                        { type: 'labelInput', id: 'tableName', labelKey: 'layout.auto.labelD31c02a0', action: `${ns}.Action.renameTable(this.value)` },
                        { type: 'large', icon: 'shengchengshujutoushibiao', labelKey: 'layout.auto.label52fa4c98', action: `${ns}.Action.resizeTable()` }
                    ]
                },
                {
                    labelKey: 'layout.auto.label9f1d3d34',
                    items: [
                        { type: 'large', icon: 'shengchengshujutoushibiao', labelKey: 'layout.auto.label765501ea', action: `${ns}.Action.createPivotTable()` },
                        { type: 'large', icon: '_chartWqiepianqi', labelKey: 'layout.auto.label694838cc', action: `${ns}.Action.insertSlicer()` },
                        {
                            type: 'stack', items: [
                                { icon: 'rongqi_2024_10_15T145712642', labelKey: 'layout.auto.labelEb15f260', action: `${ns}.Action.removeDuplicates()` },
                                { icon: 'zhuanhuan', labelKey: 'layout.auto.labelA6b39eb4', action: `${ns}.Action.convertTableToRange()` }
                            ]
                        }
                        
                    ]
                },
                {
                    labelKey: 'layout.auto.label32ac6bce',
                    items: [
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label240fc094', id: 'tableHeaderRow', checked: ctx => ctx.Layout._currentContextData?.showHeaderRow === true, action: `${ns}.Action.toggleTableOption('headerRow')` },
                                { type: 'checkbox', labelKey: 'layout.auto.labelC0185cac', id: 'tableTotalsRow', checked: ctx => ctx.Layout._currentContextData?.showTotalsRow === true, action: `${ns}.Action.toggleTableOption('totalsRow')` }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label37030c7c', id: 'tableRowStripes', checked: ctx => ctx.Layout._currentContextData?.showRowStripes === true, action: `${ns}.Action.toggleTableOption('rowStripes')` },
                                { type: 'checkbox', labelKey: 'layout.auto.label7af7db1c', id: 'tableColStripes', checked: ctx => ctx.Layout._currentContextData?.showColumnStripes === true, action: `${ns}.Action.toggleTableOption('columnStripes')` }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.labelE30beb48', id: 'tableFirstColumn', checked: ctx => ctx.Layout._currentContextData?.showFirstColumn === true, action: `${ns}.Action.toggleTableOption('firstColumn')` },
                                { type: 'checkbox', labelKey: 'layout.auto.label798a4a6c', id: 'tableLastColumn', checked: ctx => ctx.Layout._currentContextData?.showLastColumn === true, action: `${ns}.Action.toggleTableOption('lastColumn')` }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label54197afc', id: 'tableAutoFilter', checked: ctx => ctx.Layout._currentContextData?.autoFilterEnabled === true, action: `${ns}.Action.toggleTableOption('autoFilter')` }
                            ]
                        }
                    ]
                },
                {
                    labelKey: 'layout.auto.label727a825c',
                    items: [
                        { type: 'large', icon: 'biaogeyangshi', labelKey: 'layout.auto.label801897e0', hasArrow: true, menu: 'tableStyles', menuWidth: '340px' }
                    ]
                },
                {
                    labelKey: 'layout.auto.labelD932aecc',
                    items: [
                        { type: 'large', icon: 'biaogeyangshi', labelKey: 'layout.auto.label2c95d4cc', action: `${ns}.Action.deleteTable()` }
                    ]
                }
            ]
        },
        // 透视表分析面板（上下文选项卡）
        {
            key: 'pivotAnalyze',
            labelKey: 'layout.auto.labelDa7a0418',
            contextual: true,
            contextType: 'pivotTable',
            groups: [
                {
                    labelKey: 'layout.auto.label6d96702e',
                    items: [
                        { type: 'labelInput', id: 'pivotTableName', labelKey: 'layout.auto.labelD8cbc426', action: `${ns}.Action.renamePivotTable(this.value)` },
                        { type: 'large', icon: 'xuanxiang', labelKey: 'layout.auto.labelEebdf267', menu: 'pivotOptions' },
                        { type: 'labelInput', id: 'pivotTableName', labelKey: 'layout.auto.label067fc51e', action: `${ns}.Action.renamePivotTable(this.value)` },
                        { type: 'large', icon: 'xuanxiang', labelKey: 'layout.auto.labelE7274318', action: `${ns}.Action.pivotFieldSettings()` },
                        {
                            type: 'stack', items: [
                                { icon: 'niantie', labelKey: 'layout.auto.labelCa8595e4', action: `${ns}.refreshAllPivotTables()` },
                                { icon: 'niantie', labelKey: 'layout.auto.label10481acc', action: `${ns}.Action.changePivotDataSource()` }
                            ]
                        }
                    ]
                },
                {
                    labelKey: 'layout.auto.label34005110',
                    items: [
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label694838cc', action: `${ns}.Action.insertPivotSlicer()` }
                    ]
                },
                {
                    labelKey: 'layout.auto.label426d9dba',
                    items: [
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label8d9fe558', action: `${ns}.Action.refreshPivotTable()` },
                        {
                            type: 'stack', items: [
                                { icon: 'niantie', labelKey: 'layout.auto.label1b11c6bc', action: `${ns}.refreshAllPivotTables()` },
                                { icon: 'niantie', labelKey: 'layout.auto.label9371a726', action: `${ns}.Action.changePivotDataSource()` }
                            ]
                        }
                    ]
                },
                {
                    labelKey: 'layout.auto.labelD932aecc',
                    items: [
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label0b9d2bac', hasArrow: true, menu: 'pivotClear' },
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label26529d60', action: `${ns}.Action.movePivotTable()` }
                    ]
                },
                {
                    labelKey: 'layout.auto.labelF82f77ee',
                    items: [
                        { type: 'large', icon: 'niantie', id: 'pivotFieldList', labelKey: 'layout.auto.labelFd70269c', action: `${ns}.Action.togglePivotFieldList()` },
                        { type: 'large', icon: 'niantie', id: 'pivotDrillButtons', labelKey: 'layout.auto.label97beaff2', action: `${ns}.Action.togglePivotDrillButtons()` },
                        { type: 'large', icon: 'niantie', id: 'pivotFieldHeaders', labelKey: 'layout.auto.labelF43382bc', action: `${ns}.Action.togglePivotFieldHeaders()` },
                    ]
                }
            ]
        },
        // 透视表设计面板（上下文选项卡）
        {
            key: 'pivotDesign',
            labelKey: 'layout.auto.label27023994',
            contextual: true,
            contextType: 'pivotTable',
            groups: [
                {
                    labelKey: 'layout.auto.labelD4037dd0',
                    items: [
                        { type: 'large', icon: 'fenleihuizong', labelKey: 'layout.auto.label6fbb5db4', hasArrow: true, menu: 'pivotSubtotals' },
                        { type: 'large', icon: 'qiuhe', labelKey: 'layout.auto.label142c257c', hasArrow: true, menu: 'pivotGrandTotals' },
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label76fd7400', hasArrow: true, menu: 'pivotReportLayout' },
                        { type: 'large', icon: 'niantie', labelKey: 'layout.auto.label170d7ba0', hasArrow: true, menu: 'pivotBlankRows' }
                    ]
                },
                {
                    labelKey: 'layout.auto.label62572410',
                    items: [
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label9975a578', id: 'pivotRowHeaders', checked: ctx => ctx.Layout._currentContextData?.showRowHeaders !== false, action: `${ns}.Action.togglePivotOption('rowHeaders')` },
                                { type: 'checkbox', labelKey: 'layout.auto.labelCd612c50', id: 'pivotColHeaders', checked: ctx => ctx.Layout._currentContextData?.showColHeaders !== false, action: `${ns}.Action.togglePivotOption('colHeaders')` }
                            ]
                        },
                        {
                            type: 'stack', items: [
                                { type: 'checkbox', labelKey: 'layout.auto.label37030c7c', id: 'pivotRowStripes', checked: ctx => ctx.Layout._currentContextData?.showRowStripes === true, action: `${ns}.Action.togglePivotOption('rowStripes')` },
                                { type: 'checkbox', labelKey: 'layout.auto.label7af7db1c', id: 'pivotColStripes', checked: ctx => ctx.Layout._currentContextData?.showColStripes === true, action: `${ns}.Action.togglePivotOption('colStripes')` }
                            ]
                        }
                    ]
                },
                {
                    labelKey: 'layout.auto.label57267dc8',
                    items: [
                        { type: 'large', icon: 'biaogeyangshi', labelKey: 'layout.auto.label801897e0', hasArrow: true, menu: 'pivotStyles', menuWidth: '340px' }
                    ]
                }
            ]
        }
    ];

    const openConfig = _applyOpenEditionToolbarConfig(config);
    applyMinimalTextDefaults(openConfig);
    return openConfig;
}

function _applyOpenEditionToolbarConfig(config) {
    const blockedPanels = new Set(['file', 'page', 'pivotAnalyze', 'pivotDesign']);
    return config
        .filter((panel) => !blockedPanels.has(panel.key))
        .map((panel) => ({
            ...panel,
            groups: (panel.groups || [])
                .map((group) => ({
                    ...group,
                    items: _filterOpenEditionToolbarItems(group.items || [])
                }))
                .filter((group) => group.items.length > 0)
        }))
        .filter((panel) => panel.groups.length > 0);
}

function _filterOpenEditionToolbarItems(items) {
    const nextItems = [];
    for (const item of items) {
        if (_isOpenEditionBlockedItem(item)) continue;

        const next = { ...item };
        if (Array.isArray(item.items)) {
            next.items = _filterOpenEditionToolbarItems(item.items);
            if (!next.items.length && (next.type === 'row' || next.type === 'stack')) continue;
        }

        nextItems.push(next);
    }
    return nextItems;
}

function _isOpenEditionBlockedItem(item) {
    if (!item || typeof item !== 'object') return false;
    const menu = typeof item.menu === 'string' ? item.menu : '';
    const action = typeof item.action === 'string' ? item.action : '';

    const blockedMenus = new Set([
        'shapes',
        'chartBar',
        'chartLine',
        'chartPie',
        'chartArea',
        'chartScatter',
        'chartRadar',
        'pivotOptions',
        'pivotClear',
        'pivotSubtotals',
        'pivotGrandTotals',
        'pivotReportLayout',
        'pivotBlankRows',
        'pivotStyles',
        'printArea',
        'printScale',
        'pageMargins',
        'pageOrientation',
        'paperSize'
    ]);

    if (blockedMenus.has(menu)) return true;

    const blockedActionSnippets = [
        '.IO.export(',
        '.IO.import(',
        'Action.importFromUrl(',
        'Action.insertImages(',
        'Action.insertAreaScreenshot(',
        'Action.openChartModal(',
        'Action.insertChart(',
        'Action.insertTextBox(',
        '.Layout.showAIChat=',
        'Action.createPivotTable(',
        'Action.insertPivot',
        'Action.printPreview(',
        'Action.pageBreakPreview(',
        'Action.toggleShowPageBreaks(',
        'Action.insertPageBreak(',
        'Action.togglePrintGridlines(',
        'Action.togglePrintHeadings(',
        'Action.setPageMargins',
        'Action.setPageOrientation(',
        'Action.setPaperSize(',
        'Action.setPrintScale',
        'Action.setPrintAreaFromSelection(',
        'Action.clearPrintArea(',
        'Action.setPrintTitles(',
        'Action.setPageBackground(',
        'Action.setHeaderFooter('
    ];

    return blockedActionSnippets.some((snippet) => action.includes(snippet));
}


// 字体配置
const FONTS = {
    cn: ['宋体', '新宋体', '仿宋', '微软雅黑', '楷体', '黑体', '等线', '华文细黑', '华文楷体', '隶书', '幼圆'],
    en: ['Arial', 'Arial Black', 'Times New Roman', 'Verdana', 'Consolas', 'Georgia', 'Tahoma', 'Courier New', 'Calibri', 'Cambria', 'Segoe UI', 'Trebuchet MS', 'Lucida Console', 'Impact', 'Comic Sans MS']
};

function buildFontItems(ns) {
    const build = name => ({ text: name, style: `font-family:'${name}'`, action: `${ns}.Action.areasStyle('font',{name:'${name}'})` });
    return [...FONTS.cn.map(build), { type: 'divider' }, ...FONTS.en.map(build)];
}

function buildFormulaMenu(ns, categoryKey, options = {}) {
    const limit = options.limit ?? 12;
    const dialogKey = options.dialogKey ?? categoryKey;
    const formulas = FormulaCatalog.getShortcutFormulas(categoryKey, limit);
    const items = formulas.map(formula => ({
        text: formula.name,
        action: `${ns}.Action.insertFunction('${formula.name}')`
    }));
    const moreItem = { labelKey: 'layout.auto.label2b1c08f4', action: `${ns}.Action.openFormulaDialog('${dialogKey}')` };
    if (!items.length) return [moreItem];
    return [...items, { type: 'divider' }, moreItem];
}

/**
 * 创建下拉菜单配置
 */
export function createMenuConfig(ns) {
    return {
        fontName: buildFontItems(ns),
        fontSize: [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map(size => ({
            text: String(size),
            action: `${ns}.Action.areasStyle('font',{size:${size}})`
        })),
        hAlign: [
            { labelKey: 'layout.auto.label2f9e586c', icon: 'zuoduiqi', action: `${ns}.Action.areasStyle('alignment',{horizontal:'left'})` },
            { labelKey: 'layout.auto.label5957a658', icon: 'juzhongduiqi', action: `${ns}.Action.areasStyle('alignment',{horizontal:'center'})` },
            { labelKey: 'layout.auto.label09f772c4', icon: 'youduiqi', action: `${ns}.Action.areasStyle('alignment',{horizontal:'right'})` },
        ],
        vAlign: [
            { labelKey: 'layout.auto.label2d6a5a8c', icon: 'pingjunfenbuduiqiicon_huabanfuben4', action: `${ns}.Action.areasStyle('alignment',{vertical:'top'})` },
            { labelKey: 'layout.auto.labelF2b067c8', icon: 'pingjunfenbuduiqiicon_huabanfuben6', action: `${ns}.Action.areasStyle('alignment',{vertical:'center'})` },
            { labelKey: 'layout.auto.label09162f60', icon: 'pingjunfenbuduiqiicon_huabanfuben1', action: `${ns}.Action.areasStyle('alignment',{vertical:'bottom'})` },
        ],
        textRotation: [
            { labelKey: 'layout.auto.labelB1f1d5ca', action: `${ns}.Action.setTextRotation('ccw45')` },
            { labelKey: 'layout.auto.labelC1ea5002', action: `${ns}.Action.setTextRotation('cw45')` },
            { labelKey: 'layout.auto.label2e9bdfc6', action: `${ns}.Action.setTextRotation('vertical')` },
            { labelKey: 'layout.auto.label3cd54380', action: `${ns}.Action.setTextRotation('up90')` },
            { labelKey: 'layout.auto.labelCf38367c', action: `${ns}.Action.setTextRotation('down90')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label8c5bd7f4', action: `${ns}.Action.setTextRotation('none')` },
        ],
        merge: [
            { labelKey: 'layout.auto.labelDfde92a4', action: `${ns}.activeSheet.mergeCells(null,'center');${ns}._r()` },
            { labelKey: 'layout.auto.labelC2f55060', action: `${ns}.activeSheet.mergeCells();${ns}._r()` },
            { labelKey: 'layout.auto.labelD7a9dd2c', action: `${ns}.activeSheet.mergeCells(null,'same');${ns}._r()` },
            { labelKey: 'layout.auto.labelF0e14500', action: `${ns}.activeSheet.mergeCells(null,'content');${ns}._r()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label81ccfc88', action: `${ns}.activeSheet.unMergeCells();${ns}._r()` },
        ],
        numFmt: [
            { labelKey: 'toolbar.numFmt.general', action: `${ns}.Action.areasNumFmt('')` },
            { labelKey: 'toolbar.numFmt.number', extraKey: 'layout.auto.exampleD718a29c', extra: '0.01', action: `${ns}.Action.areasNumFmt('0.00')` },
            { labelKey: 'toolbar.numFmt.percent', extraKey: 'layout.auto.example2de1ff5c', extra: '0.60%', action: `${ns}.Action.areasNumFmt('0.00%')` },
            { labelKey: 'toolbar.numFmt.fraction', extraKey: 'layout.auto.example92b9c712', extra: '2 1/2', action: `${ns}.Action.areasNumFmt('# ?/?')` },
            { labelKey: 'toolbar.numFmt.scientific', extraKey: 'layout.auto.exampleCd80dac4', extra: '6.00E-03', action: `${ns}.Action.areasNumFmt('0.00E+00')` },
            { type: 'divider' },
            { labelKey: 'toolbar.numFmt.cny', extraKey: 'layout.auto.example739db61c', extra: '¥0.01', action: `${ns}.Action.areasNumFmt('¥#,##0.00')` },
            { labelKey: 'toolbar.numFmt.usd', extraKey: 'layout.auto.exampleCb87bb63', extra: '$0.01', action: `${ns}.Action.areasNumFmt('$#,##0.00')` },
            { labelKey: 'toolbar.numFmt.accounting', extraKey: 'layout.auto.example739db61c', extra: '¥0.01', action: `${ns}.Action.areasNumFmt('_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* &quot;-&quot;??_ ;_ @_ ')` },
            { type: 'divider' },
            { labelKey: 'toolbar.numFmt.date', extraKey: 'layout.auto.exampleB975896c', extra: '1900/1/1', action: `${ns}.Action.areasNumFmt('yyyy/m/d')` },
            { labelKey: 'toolbar.numFmt.longDate', extraKey: 'toolbar.numFmt.longDateExample', extra: 'January 1, 1900', action: `${ns}.Action.areasNumFmt('yyyy/m/d')` },
            { labelKey: 'toolbar.numFmt.time', extraKey: 'layout.auto.exampleB3542328', extra: '0:08:38', action: `${ns}.Action.areasNumFmt('h:mm:ss')` },
            { labelKey: 'toolbar.numFmt.dateTime', extraKey: 'layout.auto.example59160d28', extra: '2024/10/1 12:00', action: `${ns}.Action.areasNumFmt('yyyy/m/d h:mm:ss')` },
            { type: 'divider' },
            { labelKey: 'toolbar.numFmt.text', extraKey: 'layout.auto.example46c5d49a', extra: '0.006', action: `${ns}.Action.areasNumFmt('@')` },
            { type: 'divider' },
            { labelKey: 'toolbar.numFmt.custom', action: `${ns}.Action.customNumFmtDialog()` },
        ],
        currency: [
            { labelKey: 'layout.auto.label82496aef', action: `${ns}.Action.areasNumFmt('$#,##0.00')` },
            { labelKey: 'layout.auto.labelF7b3d440', action: `${ns}.Action.areasNumFmt('€#,##0.00')` },
            { labelKey: 'layout.auto.label5a640746', action: `${ns}.Action.areasNumFmt('£#,##0.00')` },
            { labelKey: 'layout.auto.labelDa8abb77', action: `${ns}.Action.areasNumFmt('¥#,##0.00')` },
            { labelKey: 'layout.auto.labelE2433858', action: `${ns}.Action.areasNumFmt('₩#,##0')` },
            { labelKey: 'layout.auto.label7844abc8', action: `${ns}.Action.areasNumFmt('₹#,##0.00')` },
            { labelKey: 'layout.auto.label47c40973', action: `${ns}.Action.areasNumFmt('₽#,##0.00')` },
            { labelKey: 'layout.auto.label4674137c', action: `${ns}.Action.areasNumFmt('CHF #,##0.00')` },
            { labelKey: 'layout.auto.label46ecf058', action: `${ns}.Action.areasNumFmt('A$#,##0.00')` },
            { labelKey: 'layout.auto.label7d40574d', action: `${ns}.Action.areasNumFmt('C$#,##0.00')` },
            { labelKey: 'layout.auto.label525c8b3c', action: `${ns}.Action.areasNumFmt('HK$#,##0.00')` },
        ],
        clear: [
            { labelKey: 'layout.auto.labelB683336c', action: `${ns}.Action.clearAreaFormat()` },
            { labelKey: 'layout.auto.labelAb4a78ac', action: `${ns}.Action.clearAreaFormat('style')` },
            { labelKey: 'layout.auto.label80c65db8', action: `${ns}.Action.clearAreaFormat('value')` },
            { labelKey: 'layout.auto.label83c5b1ec', action: `${ns}.Action.clearAreaFormat('border')` },
        ],
        protection: [
            { labelKey: 'layout.auto.labelEa7f2a4b', action: `${ns}.Action.configProtection()` },
            { labelKey: 'layout.auto.label476f4998', action: `${ns}.Action.configCellProtection()` },
            { labelKey: 'layout.auto.label6ea6ad84', action: `${ns}.Action.cancelProtection()` },
        ],
        freeze: [
            { labelKey: 'layout.auto.label06a5c768', action: `${ns}.Action.freezeAtActiveCell()`, titleKey: 'layout.auto.titleAb81d070' },
            { type: 'divider' },
            { labelKey: 'layout.auto.label12163f5a', action: `${ns}.Action.freezeRows(1)` },
            { labelKey: 'layout.auto.labelB99d153f', action: `${ns}.Action.freezeCols(1)` },
            { labelKey: 'layout.auto.labelDba82348', action: `${ns}.Action.freezeRowsAndCols(1,1)` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label8fd62c48', action: `${ns}.Action.freezeToCurrentRow()` },
            { labelKey: 'layout.auto.labelB9b1a1f8', action: `${ns}.Action.freezeToCurrentCol()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.labelD8bfe629', action: `${ns}.Action.unfreeze()` },
        ],
        saveAs: [
            { labelKey: 'layout.auto.label4868045c', action: `${ns}.IO.downloadXlsx()` },
            { labelKey: 'layout.auto.labelE62a5aa0', action: `${ns}.IO.downloadCsv()` },
            { labelKey: 'layout.auto.label5ed9f2ae', action: `${ns}.IO.downloadJson()` },
        ],
        zoom: [
            { labelKey: 'layout.auto.label20989160', action: `${ns}.activeSheet.zoom=0.5` },
            { labelKey: 'layout.auto.label28d8b828', action: `${ns}.activeSheet.zoom=0.75` },
            { labelKey: 'layout.auto.label2f97df3c', action: `${ns}.activeSheet.zoom=1` },
            { labelKey: 'layout.auto.labelE8d30c78', action: `${ns}.activeSheet.zoom=1.25` },
            { labelKey: 'layout.auto.label58e5db03', action: `${ns}.activeSheet.zoom=1.5` },
            { labelKey: 'layout.auto.label60974ee4', action: `${ns}.activeSheet.zoom=2` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label8ad96a3c', action: `${ns}.Action.zoomCustom()` },
        ],
        pageMargins: [
            { labelKey: 'layout.menu.pageMargins.normal', action: `${ns}.Action.setPageMarginsPreset('normal')` },
            { labelKey: 'layout.menu.pageMargins.narrow', action: `${ns}.Action.setPageMarginsPreset('narrow')` },
            { labelKey: 'layout.menu.pageMargins.wide', action: `${ns}.Action.setPageMarginsPreset('wide')` },
            { type: 'divider' },
            { labelKey: 'layout.menu.custom', action: `${ns}.Action.setPageMarginsDialog()` },
        ],
        pageOrientation: [
            { labelKey: 'layout.menu.pageOrientation.portrait', action: `${ns}.Action.setPageOrientation('portrait')` },
            { labelKey: 'layout.menu.pageOrientation.landscape', action: `${ns}.Action.setPageOrientation('landscape')` },
        ],
        paperSize: [
            { labelKey: 'layout.menu.paperSize.a4', action: `${ns}.Action.setPaperSize('A4')` },
            { labelKey: 'layout.menu.paperSize.letter', action: `${ns}.Action.setPaperSize('Letter')` },
        ],
        printArea: [
            { labelKey: 'layout.menu.printArea.set', action: `${ns}.Action.setPrintAreaFromSelection()` },
            { labelKey: 'layout.menu.printArea.clear', action: `${ns}.Action.clearPrintArea()` },
        ],
        printScale: [
            { labelKey: 'layout.auto.label2f97df3c', action: `${ns}.Action.setPrintScale(100)` },
            { labelKey: 'layout.auto.labelCacab6cb', action: `${ns}.Action.setPrintScale(90)` },
            { labelKey: 'layout.auto.label38a173b8', action: `${ns}.Action.setPrintScale(80)` },
            { labelKey: 'layout.auto.labelA9ccb748', action: `${ns}.Action.setPrintScale(70)` },
            { type: 'divider' },
            { labelKey: 'layout.menu.custom', action: `${ns}.Action.setPrintScaleDialog()` },
        ],
        paste: [
            { labelKey: 'layout.auto.label3d33e8d4', action: `${ns}.Action.paste('all')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label39583efc', action: `${ns}.Action.paste('value')` },
            { labelKey: 'layout.auto.label5613e7b9', action: `${ns}.Action.paste('formula')` },
            { labelKey: 'layout.auto.label9c5bc988', action: `${ns}.Action.paste('format')` },
            { labelKey: 'layout.auto.label36d826ce', action: `${ns}.Action.paste('noBorder')` },
            { labelKey: 'layout.auto.label0b046e3f', action: `${ns}.Action.paste('colWidth')` },
            { labelKey: 'layout.auto.labelE7546746', action: `${ns}.Action.paste('transpose')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label00724f4c', action: `${ns}.Action.paste('valueNumberFormat')` },
            { labelKey: 'layout.auto.label6ac6d174', action: `${ns}.Action.paste('formulaNumberFormat')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.labelD6d89414', action: `${ns}.Action.paste('comment')` },
            { labelKey: 'layout.auto.label065032d4', action: `${ns}.Action.paste('validation')` },
        ],
        convert: [
            { labelKey: 'layout.auto.label40bdc8a4', action: `${ns}.Action.textToNumber()` },
            { labelKey: 'layout.auto.labelB67479f3', action: `${ns}.Action.numberToText()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label2eedfce4', action: `${ns}.Action.toLowercase()` },
            { labelKey: 'layout.auto.label64f57004', action: `${ns}.Action.toUppercase()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label28152900', action: `${ns}.Action.trimSpaces()` },
        ],
        dataValidation: [
            { labelKey: 'layout.menu.dataValidation.open', action: `${ns}.Action.dataValidationDialog()` },
            { type: 'divider' },
            { labelKey: 'layout.menu.dataValidation.clear', action: `${ns}.Action.clearDataValidation()` }
        ],
        rowColSize: [
            { labelKey: 'layout.auto.labelE9cb6cfc', action: `${ns}.Action.setRowHeightDialog()` },
            { labelKey: 'layout.auto.labelDbcee390', action: `${ns}.Action.setColWidthDialog()` },
            { labelKey: 'layout.auto.label591efe60', action: `${ns}.Action.autoFitRowHeight()` },
            { labelKey: 'layout.auto.labelC14245a4', action: `${ns}.Action.autoFitColWidth()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.labelDda0aee0', action: `${ns}.Action.setDefaultRowHeightDialog()` },
            { labelKey: 'layout.auto.labelA0c3dddc', action: `${ns}.Action.setDefaultColWidthDialog()` },
        ],
        autoSum: [
            { labelKey: 'layout.auto.label304be598', action: `${ns}.Action.insertFunction('SUM')` },
            { labelKey: 'layout.auto.labelB191c97c', action: `${ns}.Action.insertFunction('AVERAGE')` },
            { labelKey: 'layout.auto.label0c22da14', action: `${ns}.Action.insertFunction('COUNT')` },
            { labelKey: 'layout.auto.label5d3d9fbc', action: `${ns}.Action.insertFunction('MAX')` },
            { labelKey: 'layout.auto.labelD13c02d3', action: `${ns}.Action.insertFunction('MIN')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label2b1c08f4', action: `${ns}.Action.openFormulaDialog('math')` },
        ],
        sparkline: [
            { labelKey: 'layout.auto.label8b22223e', action: `${ns}.Action.openSparklineDialog('line')` },
            { labelKey: 'layout.auto.labelE5f2b110', action: `${ns}.Action.openSparklineDialog('column')` },
            { labelKey: 'layout.auto.label637e3b1e', action: `${ns}.Action.openSparklineDialog('win_loss')` }
        ],
        textBox: [
            { labelKey: 'layout.auto.label16d80ff6', action: `${ns}.Action.insertTextBox('horizontal')` },
            { labelKey: 'layout.auto.labelB1fa2802', action: `${ns}.Action.insertTextBox('vertical')` }
        ],
        sort: [
            { labelKey: 'layout.auto.label27c10e35', action: `${ns}.Action.sortAsc()` },
            { labelKey: 'layout.auto.label56ec896c', action: `${ns}.Action.sortDesc()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label3fdb3ab0', action: `${ns}.Action.customSort()` },
        ],
        filter: [
            { labelKey: 'layout.auto.label2c508544', action: `${ns}.Action.toggleAutoFilter()` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label0fba4654', action: `${ns}.Action.clearFilter()` },
            { labelKey: 'layout.auto.label53f4d744', action: `${ns}.Action.reapplyFilter()` },
        ],
        find: [
            { labelKey: 'findReplace.tab.find', action: `${ns}.Action.openFindDialog('find')` },
            { labelKey: 'findReplace.tab.replace', action: `${ns}.Action.openFindDialog('replace')` },
            { type: 'divider' },
            { labelKey: 'findReplace.tab.goto', action: `${ns}.Action.openFindDialog('goto')` },
        ],
        removeArrows: [
            { labelKey: 'layout.auto.label7673ae70', action: `${ns}.Action.removeArrows('precedents')` },
            { labelKey: 'layout.auto.label6e89ee74', action: `${ns}.Action.removeArrows('dependents')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label7c9ea1ac', action: `${ns}.Action.removeArrows('all')` },
        ],
        calcOptions: [
            { labelKey: 'layout.auto.labelE2d8213c', action: `${ns}.Action.setCalcMode('auto')`, checked: () => window[ns]?.calcMode !== 'manual' },
            { labelKey: 'layout.auto.label6f816c60', action: `${ns}.Action.setCalcMode('manual')`, checked: () => window[ns]?.calcMode === 'manual' },
        ],
        // 表格样式菜单 - 特殊渲染
        tableStyles: 'special:tableStyles',
        // 套用表格格式菜单 - 特殊渲染
        formatAsTable: 'special:formatAsTable',
        // 单元格样式菜单 - 特殊渲染
        cellStyles: 'special:cellStyles',
        // 条件格式菜单 - 特殊渲染，这里只定义结构
        condFormat: 'special:condFormat',
        // 透视表菜单
        pivotOptions: [
            { labelKey: 'layout.auto.label8b46f723', action: `${ns}.Action.openPivotTableOptions()` },
            { labelKey: 'layout.auto.label5f1a7cbc', action: `${ns}.Action.showPivotReportFilterPages()` }
        ],
        pivotClear: [
            { labelKey: 'layout.auto.labelAac1ce9c', action: `${ns}.Action.clearPivotTable('all')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.labelFde78e74', action: `${ns}.Action.clearPivotTable('filters')` },
        ],
        pivotSubtotals: [
            { labelKey: 'layout.auto.labelB6f3db09', action: `${ns}.Action.setPivotSubtotals('none')` },
            { labelKey: 'layout.auto.labelF53301bc', action: `${ns}.Action.setPivotSubtotals('bottom')` },
            { labelKey: 'layout.auto.label45d4ed20', action: `${ns}.Action.setPivotSubtotals('top')` },
        ],
        pivotGrandTotals: [
            { labelKey: 'layout.auto.label1b148794', action: `${ns}.Action.setPivotGrandTotals('none')` },
            { labelKey: 'layout.auto.label42a69c1c', action: `${ns}.Action.setPivotGrandTotals('both')` },
            { labelKey: 'layout.auto.labelEac219e0', action: `${ns}.Action.setPivotGrandTotals('rows')` },
            { labelKey: 'layout.auto.label46acc0ac', action: `${ns}.Action.setPivotGrandTotals('cols')` },
        ],
        pivotReportLayout: [
            { labelKey: 'layout.auto.labelB290c756', action: `${ns}.Action.setPivotLayout('compact')` },
            { labelKey: 'layout.auto.label69e81956', action: `${ns}.Action.setPivotLayout('outline')` },
            { labelKey: 'layout.auto.label2364c11e', action: `${ns}.Action.setPivotLayout('tabular')` },
            { type: 'divider' },
            { labelKey: 'layout.auto.label7cc4a190', action: `${ns}.Action.setPivotRepeatLabels(true)` },
            { labelKey: 'layout.auto.label669baeb8', action: `${ns}.Action.setPivotRepeatLabels(false)` },
        ],
        pivotBlankRows: [
            { labelKey: 'layout.auto.labelB0a4de64', action: `${ns}.Action.setPivotBlankRows(true)` },
            { labelKey: 'layout.auto.labelCdcccef2', action: `${ns}.Action.setPivotBlankRows(false)` },
        ],
        // 透视表样式菜单 - 特殊渲染
        pivotStyles: 'special:pivotStyles',
        commonFn: buildFormulaMenu(ns, 'common', { limit: 12, dialogKey: 'recent' }),
        financeFn: buildFormulaMenu(ns, 'financial'),
        logicFn: buildFormulaMenu(ns, 'logical'),
        textFn: buildFormulaMenu(ns, 'text'),
        dateFn: buildFormulaMenu(ns, 'datetime'),
        lookupFn: buildFormulaMenu(ns, 'lookup'),
        mathFn: buildFormulaMenu(ns, 'math')
    };
}
