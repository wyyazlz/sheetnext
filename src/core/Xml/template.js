// ==================== XML 默认模板对象 ====================

export default {
    "[Content_Types].xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "Types": {
            "Default": [
                {
                    "_$Extension": "rels",
                    "_$ContentType": "application/vnd.openxmlformats-package.relationships+xml"
                },
                {
                    "_$Extension": "xml",
                    "_$ContentType": "application/xml"
                }
            ],
            "Override": [
                {
                    "_$PartName": "/xl/workbook.xml",
                    "_$ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                },
                {
                    "_$PartName": "/xl/worksheets/sheet1.xml",
                    "_$ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                },
                {
                    "_$PartName": "/xl/theme/theme1.xml",
                    "_$ContentType": "application/vnd.openxmlformats-officedocument.theme+xml"
                },
                {
                    "_$PartName": "/xl/styles.xml",
                    "_$ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
                },
                {
                    "_$PartName": "/docProps/core.xml",
                    "_$ContentType": "application/vnd.openxmlformats-package.core-properties+xml"
                },
                {
                    "_$PartName": "/docProps/app.xml",
                    "_$ContentType": "application/vnd.openxmlformats-officedocument.extended-properties+xml"
                }
            ],
            "_$xmlns": "http://schemas.openxmlformats.org/package/2006/content-types"
        }
    },
    "_rels/.rels": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "Relationships": {
            "Relationship": [
                {
                    "_$Id": "rId3",
                    "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                    "_$Target": "docProps/app.xml"
                },
                {
                    "_$Id": "rId2",
                    "_$Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                    "_$Target": "docProps/core.xml"
                },
                {
                    "_$Id": "rId1",
                    "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "_$Target": "xl/workbook.xml"
                }
            ],
            "_$xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
        }
    },
    "xl/workbook.xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "workbook": {
            "fileVersion": {
                "_$appName": "xl",
                "_$lastEdited": "7",
                "_$lowestEdited": "5",
                "_$rupBuild": "28025"
            },
            "workbookPr": "",
            "mc:AlternateContent": {
                "mc:Choice": {
                    "x15ac:absPath": {
                        "_$url": "C:\\Users\\29486\\Desktop\\EEEEE\\",
                        "_$xmlns:x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"
                    },
                    "_$Requires": "x15"
                },
                "_$xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
            },
            "xr:revisionPtr": {
                "_$revIDLastSave": "0",
                "_$documentId": "13_ncr:1_{A2D52201-FDF6-42D3-934A-EE0E0EC50B54}",
                "_$xr6:coauthVersionLast": "47",
                "_$xr6:coauthVersionMax": "47",
                "_$xr10:uidLastSave": "{00000000-0000-0000-0000-000000000000}"
            },
            "bookViews": {
                "workbookView": {
                    "_$xWindow": "-98",
                    "_$yWindow": "-98",
                    "_$windowWidth": "24196",
                    "_$windowHeight": "14476",
                    "_$xr2:uid": "{00000000-000D-0000-FFFF-FFFF00000000}",
                    "_$activeTab": 0
                }
            },
            "sheets": {
                "sheet": [
                    {
                        "_$name": "Sheet1",
                        "_$sheetId": "1",
                        "_$r:id": "rId1"
                    }
                ]
            },
            "calcPr": {
                "_$calcId": "144525"
            },
            "_$xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "_$xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "_$mc:Ignorable": "x15 xr xr6 xr10 xr2",
            "_$xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
            "_$xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
            "_$xmlns:xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
            "_$xmlns:xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
            "_$xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
        }
    },
    "xl/_rels/workbook.xml.rels": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "Relationships": {
            "Relationship": [
                {
                    "_$Id": "rId3",
                    "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                    "_$Target": "styles.xml"
                },
                {
                    "_$Id": "rId2",
                    "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                    "_$Target": "theme/theme1.xml"
                },
                {
                    "_$Id": "rId1",
                    "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    "_$Target": "worksheets/sheet1.xml"
                }
            ],
            "_$xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
        }
    },
    "xl/worksheets/sheet1.xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "worksheet": {
            "dimension": {
                "_$ref": "A1:Z200"
            },
            "sheetViews": {
                "sheetView": [
                    {
                        "_$tabSelected": "0",
                        "_$workbookViewId": "0"
                    }
                ]
            },
            "sheetFormatPr": {
                "_$defaultColWidth": "9.5",
                "_$defaultRowHeight": "15",
                "_$x14ac:dyDescent": "0.3"
            },
            "sheetData": "",
            "phoneticPr": {
                "_$fontId": "1",
                "_$type": "noConversion"
            },
            "pageMargins": {
                "_$left": "0.7",
                "_$right": "0.7",
                "_$top": "0.75",
                "_$bottom": "0.75",
                "_$header": "0.3",
                "_$footer": "0.3"
            },
            "pageSetup": {
                "_$paperSize": "9",
                "_$orientation": "portrait"
            },
            "_$xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "_$xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "_$mc:Ignorable": "x14ac xr xr2 xr3",
            "_$xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
            "_$xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
            "_$xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
            "_$xmlns:xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
        }
    },
    "xl/theme/theme1.xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "a:theme": {
            "a:themeElements": {
                "a:clrScheme": {
                    "a:dk1": {
                        "a:sysClr": {
                            "_$val": "windowText",
                            "_$lastClr": "000000"
                        }
                    },
                    "a:lt1": {
                        "a:sysClr": {
                            "_$val": "window",
                            "_$lastClr": "FFFFFF"
                        }
                    },
                    "a:dk2": {
                        "a:srgbClr": {
                            "_$val": "44546A"
                        }
                    },
                    "a:lt2": {
                        "a:srgbClr": {
                            "_$val": "E7E6E6"
                        }
                    },
                    "a:accent1": {
                        "a:srgbClr": {
                            "_$val": "4874CB"
                        }
                    },
                    "a:accent2": {
                        "a:srgbClr": {
                            "_$val": "EE822F"
                        }
                    },
                    "a:accent3": {
                        "a:srgbClr": {
                            "_$val": "F2BA02"
                        }
                    },
                    "a:accent4": {
                        "a:srgbClr": {
                            "_$val": "75BD42"
                        }
                    },
                    "a:accent5": {
                        "a:srgbClr": {
                            "_$val": "30C0B4"
                        }
                    },
                    "a:accent6": {
                        "a:srgbClr": {
                            "_$val": "E54C5E"
                        }
                    },
                    "a:hlink": {
                        "a:srgbClr": {
                            "_$val": "0026E5"
                        }
                    },
                    "a:folHlink": {
                        "a:srgbClr": {
                            "_$val": "7E1FAD"
                        }
                    },
                    "_$name": "WPS"
                },
                "a:fontScheme": {
                    "a:majorFont": {
                        "a:latin": {
                            "_$typeface": "Calibri Light"
                        },
                        "a:ea": {
                            "_$typeface": ""
                        },
                        "a:cs": {
                            "_$typeface": ""
                        },
                        "a:font": [
                            {
                                "_$script": "Jpan",
                                "_$typeface": "ＭＳ Ｐゴシック"
                            },
                            {
                                "_$script": "Hang",
                                "_$typeface": "맑은 고딕"
                            },
                            {
                                "_$script": "Hans",
                                "_$typeface": "宋体"
                            },
                            {
                                "_$script": "Hant",
                                "_$typeface": "新細明體"
                            },
                            {
                                "_$script": "Arab",
                                "_$typeface": "Times New Roman"
                            },
                            {
                                "_$script": "Hebr",
                                "_$typeface": "Times New Roman"
                            },
                            {
                                "_$script": "Thai",
                                "_$typeface": "Tahoma"
                            },
                            {
                                "_$script": "Ethi",
                                "_$typeface": "Nyala"
                            },
                            {
                                "_$script": "Beng",
                                "_$typeface": "Vrinda"
                            },
                            {
                                "_$script": "Gujr",
                                "_$typeface": "Shruti"
                            },
                            {
                                "_$script": "Khmr",
                                "_$typeface": "MoolBoran"
                            },
                            {
                                "_$script": "Knda",
                                "_$typeface": "Tunga"
                            },
                            {
                                "_$script": "Guru",
                                "_$typeface": "Raavi"
                            },
                            {
                                "_$script": "Cans",
                                "_$typeface": "Euphemia"
                            },
                            {
                                "_$script": "Cher",
                                "_$typeface": "Plantagenet Cherokee"
                            },
                            {
                                "_$script": "Yiii",
                                "_$typeface": "Microsoft Yi Baiti"
                            },
                            {
                                "_$script": "Tibt",
                                "_$typeface": "Microsoft Himalaya"
                            },
                            {
                                "_$script": "Thaa",
                                "_$typeface": "MV Boli"
                            },
                            {
                                "_$script": "Deva",
                                "_$typeface": "Mangal"
                            },
                            {
                                "_$script": "Telu",
                                "_$typeface": "Gautami"
                            },
                            {
                                "_$script": "Taml",
                                "_$typeface": "Latha"
                            },
                            {
                                "_$script": "Syrc",
                                "_$typeface": "Estrangelo Edessa"
                            },
                            {
                                "_$script": "Orya",
                                "_$typeface": "Kalinga"
                            },
                            {
                                "_$script": "Mlym",
                                "_$typeface": "Kartika"
                            },
                            {
                                "_$script": "Laoo",
                                "_$typeface": "DokChampa"
                            },
                            {
                                "_$script": "Sinh",
                                "_$typeface": "Iskoola Pota"
                            },
                            {
                                "_$script": "Mong",
                                "_$typeface": "Mongolian Baiti"
                            },
                            {
                                "_$script": "Viet",
                                "_$typeface": "Times New Roman"
                            },
                            {
                                "_$script": "Uigh",
                                "_$typeface": "Microsoft Uighur"
                            },
                            {
                                "_$script": "Geor",
                                "_$typeface": "Sylfaen"
                            }
                        ]
                    },
                    "a:minorFont": {
                        "a:latin": {
                            "_$typeface": "Calibri"
                        },
                        "a:ea": {
                            "_$typeface": ""
                        },
                        "a:cs": {
                            "_$typeface": ""
                        },
                        "a:font": [
                            {
                                "_$script": "Jpan",
                                "_$typeface": "ＭＳ Ｐゴシック"
                            },
                            {
                                "_$script": "Hang",
                                "_$typeface": "맑은 고딕"
                            },
                            {
                                "_$script": "Hans",
                                "_$typeface": "宋体"
                            },
                            {
                                "_$script": "Hant",
                                "_$typeface": "新細明體"
                            },
                            {
                                "_$script": "Arab",
                                "_$typeface": "Arial"
                            },
                            {
                                "_$script": "Hebr",
                                "_$typeface": "Arial"
                            },
                            {
                                "_$script": "Thai",
                                "_$typeface": "Tahoma"
                            },
                            {
                                "_$script": "Ethi",
                                "_$typeface": "Nyala"
                            },
                            {
                                "_$script": "Beng",
                                "_$typeface": "Vrinda"
                            },
                            {
                                "_$script": "Gujr",
                                "_$typeface": "Shruti"
                            },
                            {
                                "_$script": "Khmr",
                                "_$typeface": "DaunPenh"
                            },
                            {
                                "_$script": "Knda",
                                "_$typeface": "Tunga"
                            },
                            {
                                "_$script": "Guru",
                                "_$typeface": "Raavi"
                            },
                            {
                                "_$script": "Cans",
                                "_$typeface": "Euphemia"
                            },
                            {
                                "_$script": "Cher",
                                "_$typeface": "Plantagenet Cherokee"
                            },
                            {
                                "_$script": "Yiii",
                                "_$typeface": "Microsoft Yi Baiti"
                            },
                            {
                                "_$script": "Tibt",
                                "_$typeface": "Microsoft Himalaya"
                            },
                            {
                                "_$script": "Thaa",
                                "_$typeface": "MV Boli"
                            },
                            {
                                "_$script": "Deva",
                                "_$typeface": "Mangal"
                            },
                            {
                                "_$script": "Telu",
                                "_$typeface": "Gautami"
                            },
                            {
                                "_$script": "Taml",
                                "_$typeface": "Latha"
                            },
                            {
                                "_$script": "Syrc",
                                "_$typeface": "Estrangelo Edessa"
                            },
                            {
                                "_$script": "Orya",
                                "_$typeface": "Kalinga"
                            },
                            {
                                "_$script": "Mlym",
                                "_$typeface": "Kartika"
                            },
                            {
                                "_$script": "Laoo",
                                "_$typeface": "DokChampa"
                            },
                            {
                                "_$script": "Sinh",
                                "_$typeface": "Iskoola Pota"
                            },
                            {
                                "_$script": "Mong",
                                "_$typeface": "Mongolian Baiti"
                            },
                            {
                                "_$script": "Viet",
                                "_$typeface": "Arial"
                            },
                            {
                                "_$script": "Uigh",
                                "_$typeface": "Microsoft Uighur"
                            },
                            {
                                "_$script": "Geor",
                                "_$typeface": "Sylfaen"
                            }
                        ]
                    },
                    "_$name": "WPS"
                },
                "a:fmtScheme": {
                    "a:fillStyleLst": {
                        "a:solidFill": {
                            "a:schemeClr": {
                                "_$val": "phClr"
                            }
                        },
                        "a:gradFill": [
                            {
                                "a:gsLst": {
                                    "a:gs": [
                                        {
                                            "a:schemeClr": {
                                                "a:lumOff": {
                                                    "_$val": "17500"
                                                },
                                                "_$val": "phClr"
                                            },
                                            "_$pos": "0"
                                        },
                                        {
                                            "a:schemeClr": {
                                                "_$val": "phClr"
                                            },
                                            "_$pos": "100000"
                                        }
                                    ]
                                },
                                "a:lin": {
                                    "_$ang": "2700000",
                                    "_$scaled": "0"
                                }
                            },
                            {
                                "a:gsLst": {
                                    "a:gs": [
                                        {
                                            "a:schemeClr": {
                                                "a:hueOff": {
                                                    "_$val": "-2520000"
                                                },
                                                "_$val": "phClr"
                                            },
                                            "_$pos": "0"
                                        },
                                        {
                                            "a:schemeClr": {
                                                "_$val": "phClr"
                                            },
                                            "_$pos": "100000"
                                        }
                                    ]
                                },
                                "a:lin": {
                                    "_$ang": "2700000",
                                    "_$scaled": "0"
                                }
                            }
                        ]
                    },
                    "a:lnStyleLst": {
                        "a:ln": [
                            {
                                "a:solidFill": {
                                    "a:schemeClr": {
                                        "_$val": "phClr"
                                    }
                                },
                                "a:prstDash": {
                                    "_$val": "solid"
                                },
                                "a:miter": {
                                    "_$lim": "800000"
                                },
                                "_$w": "12700",
                                "_$cap": "flat",
                                "_$cmpd": "sng",
                                "_$algn": "ctr"
                            },
                            {
                                "a:solidFill": {
                                    "a:schemeClr": {
                                        "_$val": "phClr"
                                    }
                                },
                                "a:prstDash": {
                                    "_$val": "solid"
                                },
                                "a:miter": {
                                    "_$lim": "800000"
                                },
                                "_$w": "12700",
                                "_$cap": "flat",
                                "_$cmpd": "sng",
                                "_$algn": "ctr"
                            },
                            {
                                "a:gradFill": {
                                    "a:gsLst": {
                                        "a:gs": [
                                            {
                                                "a:schemeClr": {
                                                    "a:hueOff": {
                                                        "_$val": "-4200000"
                                                    },
                                                    "_$val": "phClr"
                                                },
                                                "_$pos": "0"
                                            },
                                            {
                                                "a:schemeClr": {
                                                    "_$val": "phClr"
                                                },
                                                "_$pos": "100000"
                                            }
                                        ]
                                    },
                                    "a:lin": {
                                        "_$ang": "2700000",
                                        "_$scaled": "1"
                                    }
                                },
                                "a:prstDash": {
                                    "_$val": "solid"
                                },
                                "a:miter": {
                                    "_$lim": "800000"
                                },
                                "_$w": "12700",
                                "_$cap": "flat",
                                "_$cmpd": "sng",
                                "_$algn": "ctr"
                            }
                        ]
                    },
                    "a:effectStyleLst": {
                        "a:effectStyle": [
                            {
                                "a:effectLst": {
                                    "a:outerShdw": {
                                        "a:schemeClr": {
                                            "a:alpha": {
                                                "_$val": "60000"
                                            },
                                            "_$val": "phClr"
                                        },
                                        "_$blurRad": "101600",
                                        "_$dist": "50800",
                                        "_$dir": "5400000",
                                        "_$algn": "ctr",
                                        "_$rotWithShape": "0"
                                    }
                                }
                            },
                            {
                                "a:effectLst": {
                                    "a:reflection": {
                                        "_$stA": "50000",
                                        "_$endA": "300",
                                        "_$endPos": "40000",
                                        "_$dist": "25400",
                                        "_$dir": "5400000",
                                        "_$sy": "-100000",
                                        "_$algn": "bl",
                                        "_$rotWithShape": "0"
                                    }
                                }
                            },
                            {
                                "a:effectLst": {
                                    "a:outerShdw": {
                                        "a:srgbClr": {
                                            "a:alpha": {
                                                "_$val": "63000"
                                            },
                                            "_$val": "000000"
                                        },
                                        "_$blurRad": "57150",
                                        "_$dist": "19050",
                                        "_$dir": "5400000",
                                        "_$algn": "ctr",
                                        "_$rotWithShape": "0"
                                    }
                                }
                            }
                        ]
                    },
                    "a:bgFillStyleLst": {
                        "a:solidFill": [
                            {
                                "a:schemeClr": {
                                    "_$val": "phClr"
                                }
                            },
                            {
                                "a:schemeClr": {
                                    "a:tint": {
                                        "_$val": "95000"
                                    },
                                    "a:satMod": {
                                        "_$val": "170000"
                                    },
                                    "_$val": "phClr"
                                }
                            }
                        ],
                        "a:gradFill": {
                            "a:gsLst": {
                                "a:gs": [
                                    {
                                        "a:schemeClr": {
                                            "a:tint": {
                                                "_$val": "93000"
                                            },
                                            "a:satMod": {
                                                "_$val": "150000"
                                            },
                                            "a:shade": {
                                                "_$val": "98000"
                                            },
                                            "a:lumMod": {
                                                "_$val": "102000"
                                            },
                                            "_$val": "phClr"
                                        },
                                        "_$pos": "0"
                                    },
                                    {
                                        "a:schemeClr": {
                                            "a:tint": {
                                                "_$val": "98000"
                                            },
                                            "a:satMod": {
                                                "_$val": "130000"
                                            },
                                            "a:shade": {
                                                "_$val": "90000"
                                            },
                                            "a:lumMod": {
                                                "_$val": "103000"
                                            },
                                            "_$val": "phClr"
                                        },
                                        "_$pos": "50000"
                                    },
                                    {
                                        "a:schemeClr": {
                                            "a:shade": {
                                                "_$val": "63000"
                                            },
                                            "a:satMod": {
                                                "_$val": "120000"
                                            },
                                            "_$val": "phClr"
                                        },
                                        "_$pos": "100000"
                                    }
                                ]
                            },
                            "a:lin": {
                                "_$ang": "5400000",
                                "_$scaled": "0"
                            },
                            "_$rotWithShape": "1"
                        }
                    },
                    "_$name": "WPS"
                }
            },
            "a:objectDefaults": "",
            "a:extraClrSchemeLst": "",
            "_$xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "_$name": "WPS"
        }
    },
    "xl/styles.xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "styleSheet": {
            "numFmts": {
                "numFmt": [],
                "_$count": "0",
            },
            "fonts": {
                "font": [
                    {
                        "sz": {
                            "_$val": "11"
                        },
                        "color": {
                            "_$theme": "1"
                        },
                        "name": {
                            "_$val": "宋体"
                        },
                        "charset": {
                            "_$val": "134"
                        },
                        "scheme": {
                            "_$val": "minor"
                        }
                    },
                    {
                        "sz": {
                            "_$val": "9"
                        },
                        "name": {
                            "_$val": "宋体"
                        },
                        "charset": {
                            "_$val": "134"
                        },
                        "scheme": {
                            "_$val": "minor"
                        }
                    }
                ],
                "_$count": "2",
                "_$x14ac:knownFonts": "1"
            },
            "fills": {
                "fill": [
                    {
                        "patternFill": {
                            "_$patternType": "none"
                        }
                    },
                    {
                        "patternFill": {
                            "_$patternType": "gray125"
                        }
                    }
                ],
                "_$count": "2"
            },
            "borders": {
                "border": [{
                    "left": "",
                    "right": "",
                    "top": "",
                    "bottom": "",
                    "diagonal": ""
                }],
                "_$count": "1"
            },
            "cellStyleXfs": {
                "xf": {
                    "alignment": {
                        "_$vertical": "center"
                    },
                    "_$numFmtId": "0",
                    "_$fontId": "0",
                    "_$fillId": "0",
                    "_$borderId": "0"
                },
                "_$count": "1"
            },
            "cellXfs": {
                "xf": [{
                    "alignment": {
                        "_$vertical": "center"
                    },
                    "_$numFmtId": "0",
                    "_$fontId": "0",
                    "_$fillId": "0",
                    "_$borderId": "0",
                    "_$xfId": "0"
                }],
                "_$count": "1"
            },
            "cellStyles": {
                "cellStyle": {
                    "_$name": "常规",
                    "_$xfId": "0",
                    "_$builtinId": "0"
                },
                "_$count": "1"
            },
            "dxfs": {
                "_$count": "0"
            },
            "tableStyles": {
                "_$count": "0",
                "_$defaultTableStyle": "TableStyleMedium2",
                "_$defaultPivotStyle": "PivotStyleLight16"
            },
            "extLst": {
                "ext": [
                    {
                        "x14:slicerStyles": {
                            "_$defaultSlicerStyle": "SlicerStyleLight1"
                        },
                        "_$uri": "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
                        "_$xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                    },
                    {
                        "x15:timelineStyles": {
                            "_$defaultTimelineStyle": "TimeSlicerStyleLight1"
                        },
                        "_$uri": "{9260A510-F301-46a8-8635-F512D64BE5F5}",
                        "_$xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                    }
                ]
            },
            "_$xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "_$xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "_$mc:Ignorable": "x14ac x16r2 xr",
            "_$xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
            "_$xmlns:x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
            "_$xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
        }
    },
    "docProps/core.xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "cp:coreProperties": {
            "dc:creator": "SheetNext Editer",
            "cp:lastModifiedBy": "SheetNext Editer",
            "dcterms:created": {
                "#text": "2023-05-12T11:15:00Z",
                "_$xsi:type": "dcterms:W3CDTF"
            },
            "dcterms:modified": {
                "#text": "2024-10-08T09:05:37Z",
                "_$xsi:type": "dcterms:W3CDTF"
            },
            "_$xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "_$xmlns:dc": "http://purl.org/dc/elements/1.1/",
            "_$xmlns:dcterms": "http://purl.org/dc/terms/",
            "_$xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
            "_$xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
        }
    },
    "docProps/app.xml": {
        "?xml": {
            "_$version": "1.0",
            "_$encoding": "UTF-8",
            "_$standalone": "yes"
        },
        "Properties": {
            "Application": "Microsoft Excel",
            "DocSecurity": "0",
            "ScaleCrop": "false",
            "HeadingPairs": {
                "vt:vector": {
                    "vt:variant": [
                        {
                            "vt:lpstr": "工作表"
                        },
                        {
                            "vt:i4": "1"
                        }
                    ],
                    "_$size": "2",
                    "_$baseType": "variant"
                }
            },
            "TitlesOfParts": {
                "vt:vector": {
                    "vt:lpstr": "Sheet1",
                    "_$size": "1",
                    "_$baseType": "lpstr"
                }
            },
            "LinksUpToDate": "false",
            "SharedDoc": "false",
            "HyperlinksChanged": "false",
            "AppVersion": "16.0300",
            "_$xmlns": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            "_$xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        }
    }
}
