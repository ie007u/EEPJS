/**
 * Name：EEP(ExcelExportPro)
 * 版本：v 1.1
 * 时间：2023年10月8日 9:51:00
 * =======================================
 * ---**基于EEP初版的重构版**---
 * 基于/resource/exceljs/xlsx.js v3.8.2
 * 官方文档： https://github.com/exceljs/exceljs/blob/master/README_zh.md
 * 该类还引用了 /js/jquery-1.8.3.mini.js
 * 使用时请务必引入以上2个js文件！
 * =======================================
 **/
(function (window, undefined) {
    var EEP = {
        /**
         * 导出配置
         **/
        config: {
            /*下载文件名*/
            name: "文件名",
            /*导出的文件后缀*/
            suffix: ".xlsx",
            /*Sheet名*/
            sheet: "Sheet",
            /*单个excel最大承容纳数据量*/
            limit: 10000,
            /*导出后直接下载*/
            isDown: true,
            /*导出字段 对应自定义导出列格式 里面存放的格式为[field_format,field_format2...]*/
            exportField: [],
            /*表头占用行数*/
            headSize: 1,
            /*冻结行数/*0为不冻结*/
            freezeSize: 0,
            /*冻结的x轴 要仅冻结行，请将其设置为 0 或 undefined*/
            freezeX: 0,
            /*冻结的y轴 要仅冻结列，请将其设置为 0 或 undefined*/
            freezeY: 0,
            /*是否启用水印*/
            isWatermark: false,
            /*水印源 可以是base64 或者 文件访问路径*/
            wateRmark: "",
            /*图片格式 png jpg..*/
            wateRmarkType: "",
            /*是否启用自动筛选*/
            isAutoScreen: false,
            /*筛选起始位置 例如 A2 代表第一列第二行*/
            screenFrom: "",
            /*筛选截止位置 不填 全部字段都自动筛选*/
            screenTo: "",
            /*是否包含合并*/
            isMerge: false,
            /*需合并的返回字段 ["1","..."]*/
            mergeFiled: [],
            /*影响合并的字 里面存放的格式为[["字段1","字段2",...],[...]]*/
            mergeFiledrule: [],
            /*是否横向延伸*/
            xExtend: false,
            /*横向字段*/
            xExtendFiled: "",
            /*横向前缀 请确保该前缀唯一不会与其他字段重复！*/
            xExtendPrefix: "",
            /*是否返回数据*/
            isresult: false,
            /*临时 url样式设置 格式 A|B 或者A */
            urltag: "",
            /*导出进度*/
            /*默认全局样式*/
            defaultStyle: {
                /*头部列样式*/
                headcolfont: {},
                /*头部行样式*/
                headrowfill: {},
                /*列样式*/
                bodycolfont: {},
                /*内容行样式*/
                bodyrowfill: {},
                /*头部行高*/
                headrowsHeight: 22,
                /*内容行高*/
                bodyrowsHeight: 22.5,
                /*链接样式*/
                linkStyle: {},
            }
        },
        /*模型*/
        model: {
            /**样式**/
            style: {
                /*(列)字体样式*/
                font: {
                    /*字体大小*/
                    size: 10,
                    /*备用字体家族。整数值。	1 - Serif, 2 - Sans Serif, 3 - Mono, Others - unknown*/
                    family: 1,
                    /*字体方案 'minor', 'major', 'none'*/
                    scheme: '',
                    /*字体例如 0000FF*/
                    color: { argb: '' },
                    /*粗细体*/
                    bold: false,
                    /*粗细*/
                    italic: '',
                    /*删除线*/
                    strike: false,
                    /*字体轮廓*/
                    outline: false,
                    /*垂直对齐*/
                    vertAlign: '',
                    /*颜色*/
                    color: { argb: '' },
                    /*字体名*/
                    name: '',
                    /*是否启用下划线*/
                    underline: false
                },
                /*(行)充满样式*/
                fill: {
                    type: 'pattern',
                    pattern: '',
                    /*背景色*/
                    fgColor: { argb: '' }
                },
                /*对齐样式*/
                alignment: {
                    /*垂直*/
                    vertical: 'top',
                    /*水平*/
                    horizontal: 'left'
                },
                /*边框 thin 单边框*/
                border: {
                    top: { style: '' },
                    left: { style: '' },
                    bottom: { style: '' },
                    right: { style: '' }
                },
                /*边框可选样式
                 *thin
                 *dotted
                 *dashDot
                 *hair
                 *dashDotDot
                 *slantDashDot
                 *mediumDashed
                 *mediumDashDotDot
                 *mediumDashDot
                 *medium
                 *double
                 *thick
                 */
                borderNode: {
                    style: 'double', color: { argb: 'FF00FF00' }
                }
            },
            /**导出字段调用格式**/
            field: { _alias: "别名/对应接口返回数据字段", _id: "如果涉及自定义导出字段此处需要使用同一字段", _name: "导出字段名", _show: 0, _width: 5, _vertical: 'middle', _horizontal: 'center' }
        },
        /*工具*/
        tools: {
            /*获取当前时间*/
            getDate: function (type = 0) {
                var myDate = new Date().toLocaleString();
                var week = new Date().getDay();
                var weeks = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
                return myDate;
            },
            /**
             * 解析url
             * @param {any} data 数据
             * @param {any} spstr 截取的字符 默认为：||
            */
            geturl: function (data, spstr = "||") {
                var str = data;
                var result = "";
                if (str && typeof str === 'string') {
                    var arr = str.split(spstr);
                    if (arr.length >= 2) {
                        var data1 = arr[0];
                        var data2 = arr[1];
                        if (data2) {
                            result = {
                                text: data1,/*内容*/
                                hyperlink: data2,/*关联链接*/
                                tooltip: data2/*提示信息*/
                            };
                        } else {
                            result = data1;
                        }
                    } else if (/^(http)/.test(arr[0])) {
                        result = {
                            text: arr[0],
                            hyperlink: arr[0],
                            tooltip: arr[0]
                        };
                    }
                    else {
                        result = data;
                    }
                }
                return result ? result : str;
            },
            /**
             * 处理最大信息长度
             * @param {any} str
            */
            splitstr: function (str) {
                if (str.length > 32767) {
                    return str.slice(0, 32767)
                } else {
                    return str;
                }
            },
            /**
             * 跳转到某个元素支持该方法的浏览器有 IE、Firefox、Safari和Opera
             * @param {any} element 要跳转的元素
             * @param {any} position 跳转位置 true代表元素头 false 元素尾
             */
            goHere: function (element, position = true) {
                if (element)
                    document.querySelector(element).scrollIntoView(position);
            },
            /**
             * 是否为允许的后缀
             * @param {any} str
             */
            isAllowSuffix: function (str) {
                if (str.indexOf(".") == -1)
                    str = "." + str;
                var rules = [".xls", ".xlsx", ".xlsm", ".xlsb", ".xltm", ".xltx"];
                return (rules.indexOf(str) != -1);
            },
            /**
             * 生成随机串
             * @param {any} length 生成长度 默认6
             */
            creatGuid: function getRandomCode(length = 6) {
                if (length > 0) {
                    var data = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"];
                    var nums = "";
                    for (var i = 0; i < length; i++) {
                        var r = parseInt(Math.random() * 61);
                        nums += data[r];
                    }
                    return nums;
                } else {
                    return false;
                }
            }
        },
        /**
        * 添加导出字段(单个加)
        * @param {any} _alias 别名\对应接口返回数据字段 必须唯一
        * @param {any} _id 元素id 不需要自定义导出可以不填
        * @param {any} _name 导出列名
        * @param {any} _show 0显示 1不显示
        * @param {any} _width 宽度 默认5 可以是float类型
        * @param {any} _vertical 垂直 默认：middle top
        * @param {any} _horizontal 水平 默认：center left  right
        */
        addField: function (_alias = '', _id = '', _name = '', _show = 0, _width = 5, _vertical = 'middle', _horizontal = 'center') {
            if (_name) {
                EEP.config.exportField.push({ alias: _alias, id: _id, name: _name, show: _show, style: { key: _alias, width: _width, alignment: { vertical: _vertical, horizontal: _horizontal } } });
            } else {
                console.log("调用addField失败！- alias或name不能为空");
            }
        },
        /**
        * 添加导出字段(批量加) 格式 [{_alias:"", _id:"", _name:"", _show : 0, _width : 5, _vertical : 'middle', _horizontal : 'center'}...]
        * @param {any} list {_alias:"", _id:"", _name:"", _show : 0, _width : 5, _vertical : 'middle', _horizontal : 'center'}
        */
        addFieldList: function (_list = []) {
            if (_list.length > 0) {
                for (var item in _list) {
                    try {
                        EEP.addField(item._alias, item._id, item._name, item._show, item._width, item._vertical, item._horizontal);
                    } catch (e) {
                        console.log(e);
                    }
                }
            }
        },
        /**
         * 导出
         * @param {any} config 导出配置
         * @param {any} content 导出数据源
         * @param {any} result 用于接收返回参数
         * @param {any} _callback 完成事件/完成后要执行的事件
         * @param {any} exportProgres 导出进度
        */
        export: function (config = EEP.config, content = [], result = [], _callback) {
            if (content.length <= 0) {
                alert("EEPJS:导出数据不能为空！");
            } else if (!config) {
                alert("EEPJS:导出配置不能为空！");
            } else if (config.exportField.length <= 0) {
                alert(msg = "EEPJS:导出字段不能为空！");
            } else if (!EEP.tools.isAllowSuffix(config.suffix)) {
                alert("导出格式错误！【" + config.suffix + "】");
            } else if (config.isMerge && config.mergeFiled.length <= 0) {
                alert("设置合并后，配置中的mergeFiled字段不能为空！ ");
            } else {
                /*数据分块*/
                var chunkSize = config.limit;
                /*确认分块量*/
                var chunkCount = Math.ceil(content.length / chunkSize);
                /*导出的总数据量*/
                var excelfiletotal = 0;
                console.log("EEPJS:本次预计导出" + chunkCount + "个文件");
                /*导出列*/
                var field_export = [];
                /*导出列控制*/
                var columns_style = [];
                /*接收的数据字段*/
                var excel_field = [];

                /*获取表头与列样式*/
                $.each(config.exportField, function (k, v) {
                    if (v.show == 0) {
                        /*拼接行数据*/
                        excel_field.push(v.alias);
                        /*导出列头*/
                        field_export.push(v.name);
                        /*导出列样式*/
                        columns_style.push(v.style);
                    }
                });

                var expid = 0;
                /* 按照分块导出数据*/
                for (var i = 0; i < chunkCount; i++) {
                    var startIndex = i * chunkSize;
                    var endIndex = Math.min(startIndex + chunkSize, content.length);
                    var dt = content.slice(startIndex, endIndex);
                    var excelmanage = [];
                    expid++;
                    /*导出名*/
                    var export_name = config.name + "_" + expid;
                    /*仅在下载的时候才会拼接数据*/
                    if (config.isDown) {
                        /*处理数据*/
                        var xyinfo = [];
                        var y2 = config.headSize + 1;
                        $.each(dt, function (k, v) {
                            var excelinfo = [];
                            var fdlist = config.mergeFiled[0];
                            /*是否有影响合并字段*/
                            if (config.isMerge && v[fdlist].length > 1) {
                                for (var i = 0; i < v[fdlist].length; i++) {
                                    excelinfo = [];
                                    if (i > 0) {
                                        $.each(excel_field, function (k, str) {
                                            if (!config.xExtend) {
                                                if ($.inArray(str, config.mergeFiledrule) > -1) {
                                                    excelinfo.push(v[fdlist][i][str])
                                                } else
                                                    excelinfo.push("")
                                            } else {
                                                if (str.indexOf(config.xExtendPrefix) == -1) {
                                                    excelinfo.push(v[fdlist][i][str]);
                                                } else {
                                                    var acontent = v[config.xExtendFiled][str.replace(config.xExtendPrefix, "")];
                                                    if (acontent) {
                                                        excelinfo.push(acontent);
                                                    } else {
                                                        excelinfo.push("");
                                                    }
                                                }
                                            }
                                        });
                                    } else {
                                        $.each(excel_field, function (k, str) {
                                            if ($.inArray(str, config.mergeFiledrule) > -1) {
                                                excelinfo.push(v[fdlist][i][str])
                                            } else
                                                excelinfo.push(EEP.tools.geturl(v[str]))
                                        });
                                    }
                                    excelmanage.push(excelinfo);
                                }
                                var y1 = y2;
                                $.each(excel_field, function (k, str) {
                                    if ($.inArray(str, config.mergeFiledrule) == -1) {
                                        var x = excel_field.indexOf(str) + 1;
                                        y2 = y1 + v[fdlist].length;
                                        xyinfo.push({ "y1": y1, "x1": x, "y2": y2 - 1, "x2": x })
                                    }
                                });
                            } else {
                                y2++;
                                $.each(excel_field, function (sy, str) {
                                    if (!config.xExtend) {
                                        excelinfo.push(EEP.tools.geturl(v[str]));
                                    } else {
                                        if (str.indexOf(config.xExtendPrefix) == -1) {
                                            excelinfo.push(EEP.tools.geturl(v[str]));
                                        } else {
                                            var acontent = v[config.xExtendFiled][str.replace(config.xExtendPrefix, "")];
                                            if (acontent) {
                                                excelinfo.push(acontent);
                                            } else {
                                                excelinfo.push("");
                                            }
                                        }
                                    }
                                });
                                excelmanage.push(excelinfo);
                            }
                        });
                        /*excel函数引用*/
                        var workbook = new ExcelJS.Workbook();

                        var sheetconfig = {};

                        if (config.freezeSize > 0) {
                            sheetconfig = {
                                views: [{ state: 'frozen', xSplit: config.freezeX, ySplit: config.freezeY }]
                            }
                        }

                        /*创建表格 并冻结前1行*/
                        var worksheet = workbook.addWorksheet(config.sheet, sheetconfig);

                        /*添加水印*/
                        if (config.isWatermark) {
                            /*添加图片 通过 base64  将图像添加到工作簿*/
                            const myBase64Imagebackground = config.wateRmark;
                            const imageId = workbook.addImage({
                                base64: myBase64Imagebackground,
                                extension: config.wateRmarkType,
                            });
                            /*添加水印背景*/
                            worksheet.addBackgroundImage(imageId);
                        }

                        /*设置字段列*/
                        var headerRow = worksheet.addRow(field_export);
                        headerRow.height = config.defaultStyle.headrowsHeight;
                        /*样式*/
                        if (config.defaultStyle.headrowfill) {
                            headerRow.fill = config.defaultStyle.headrowfill;
                        }
                        /*定义列属性*/
                        worksheet.columns = columns_style;

                        /*自动筛选*/
                        if (config.isAutoScreen) {
                            worksheet.autoFilter = {
                                from: config.screenFrom,
                                to: {
                                    row: config.freezeSize,
                                    column: (config.screenTo ? config.screenTo : field_export.length),
                                }
                            }
                        }

                        /*设置每一列的宽度 样式高度 待完善 合并到了excel_thead 的style中*/
                        for (let col = 1; col <= worksheet.columns.length; col++) {
                            const column = worksheet.getColumn(col);
                            var st_col = columns_style.filter(i => i.key == column._key);
                            if (st_col) {
                                column.alignment = ($.isEmptyObject(st_col[0].alignment) ? { vertical: 'middle', horizontal: 'center' } : st_col[0].alignment);
                            }
                        }

                        /*添加数据*/
                        $.each(excelmanage, function (k, v) {
                            var row = worksheet.addRow(v);
                            /*设置行高*/
                            row.height = config.defaultStyle.bodyrowsHeight;
                            /*设置单元字体大小*/
                            var regexPattern = new RegExp('^(' + config.urltag + ')');
                            row.eachCell(function (cell, colNumber) {
                                if (regexPattern.test(cell.address)) {
                                    cell.font = {
                                        color: { argb: '0000FF' },
                                        size: 10,
                                        underline: true
                                    };
                                } else {
                                    cell.font = { size: 10 }; /* 设置字体大小为10*/
                                }
                            });
                        });

                        /*合并单元格*/
                        $.each(xyinfo, function (k, v) {
                            worksheet.mergeCells(v.y1, v.x1, v.y2, v.x2);
                        });

                        /*用来放在页面导出位置的内容 此处的expid*/
                        if (config.name.indexOf("_") > -1) {
                            export_name = config.name;
                        }
                        /*导出时间*/
                        var export_time = EEP.tools.getDate();
                        /*导出数量*/
                        var export_total = dt.length;
                        /*验证后缀*/
                        if (config.suffix.indexOf(".") == -1)
                            config.suffix = "." + config.suffix;
                        /*返回信息*/
                        if (config.isresult) {
                            result.push({ name: export_name, sheet: config.sheet, time: export_time, total: export_total, suffix: config.suffix, info: dt });
                        }
                        /*生成并下载*/
                        workbook.xlsx.writeBuffer().then(function (buffer) {
                            excelfiletotal = excelfiletotal + dt.length;
                            var blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                            var link = document.createElement("a");
                            link.href = window.URL.createObjectURL(blob);
                            link.download = export_name + config.suffix;
                            link.click();

                            if (_callback) {
                                _callback();
                            }
                        });
                    } else {
                        result.push({ name: export_name, sheet: config.sheet, time: EEP.tools.getDate(), total: dt.length, suffix: config.suffix, info: dt });
                        /*回调事件*/
                        if (_callback) {
                            _callback();
                        }
                    }
                }
            }
        }
    }
    /*将 EPP 暴露给全局作用域*/
    window.EEP = EEP;
})(window);
