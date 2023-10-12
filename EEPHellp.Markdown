@[TOC](EEPJS-前端导出Excel的Script库)

# EEPJS简介
基于[**xlsx.js**](https://github.com/exceljs/exceljs/blob/master/README_zh.md) v3.8.2
该库还引用了**jquery-1.8.3.mini.js**

EEP JavaScript 库是用于将数据导出到 Excel 文件的工具。它提供了各种格式化选项，可以根据需要进行灵活和自定义的数据导出。
使用时请务必引入以上**2**个js文件！

## 开发起因
 
        博主本是后端开发在公司需求催化下 接触到了前端excel导出
        情景:唉 服务器负载又满了 想办法把一部分业务拆分一下吧！
        我：额 en 阿巴阿巴...
        找来找去 总算找到一款比较合适(免费且功能完善)的导出插件
        原版xlsx.js 在使用时要注意很多东西配置 就相当于是 它是直接提供基础材料
        给你,你根据需要添砖加瓦,由于导出经常用且不是一个地方调用 是很多地方调用 所以为了方便后期使用 需要对其进行封装，
        前端大佬有好建议欢迎提出 我会根据情况采纳意见~
        因为我的工作场景是后端 前端只了解基本使用,发出来也是为了 提供给同样有需要导出excel表格的
        同僚借鉴思路,也可以直接拿来使用,再就是学无止境,向愿意分享的大佬们学习·致敬
## 配置选项
EEP 库的模型包括样式、导出字段调用格式等
通过**EEP.config**调用
```javascript
{
            /* 下载文件名 */
            name: "文件名",
            /* 导出的文件后缀 */
            suffix: ".xlsx",
            /* Sheet 名 */
            sheet: "Sheet",
            /* 单个 Excel 最大承载数据量 */
            limit: 10000,
            /* 导出后直接下载 */
            isDown: true,
            /* 导出字段 对应自定义导出列格式 里面存放的格式为 [field_format, field_format2...] */
            exportField: [],
            /* 表头占用行数 */
            headSize: 1,
            /* 冻结行数 (0 为不冻结) */
            freezeSize: 0,
            /* 冻结的 x 轴 要仅冻结行，请将其设置为 0 或 undefined */
            freezeX: 0,
            /* 冻结的 y 轴 要仅冻结列，请将其设置为 0 或 undefined */
            freezeY: 0,
            /* 是否启用水印 */
            isWatermark: false,
            /* 水印源，可以是 base64 或者文件访问路径 */
            wateRmark: "",
            /* 图片格式 png jpg... */
            wateRmarkType: "",
            /* 是否启用自动筛选 */
            isAutoScreen: false,
            /* 筛选起始位置 例如 A2 代表第一列第二行 */
            screenFrom: "",
            /* 筛选截止位置 不填全部字段都自动筛选 */
            screenTo: "",
            /* 是否包含合并 */
            isMerge: false,
            /* 需合并的返回字段 ["1","..."] */
            mergeFiled: [],
            /* 影响合并的字 里面存放的格式为 [["字段1", "字段2", ...], [...]] */
            mergeFiledrule: [],
            /* 是否横向延伸 */
            xExtend: false,
            /* 横向字段 */
            xExtendFiled: "",
            /* 横向前缀，请确保该前缀唯一不会与其他字段重复 */
            xExtendPrefix: "",
            /* 是否返回数据 */
            isresult: false,
            /* 临时 url 样式设置，格式 A|B 或者 A */
            urltag: "",
            /* 导出进度 */
            /* 默认全局样式 */
            defaultStyle: {
                /* 头部列样式 */
                headcolfont: {},
                /* 头部行样式 */
                headrowfill: {},
                /* 列样式 */
                bodycolfont: {},
                /* 内容行样式 */
                bodyrowfill: {},
                /* 头部行高 */
                headrowsHeight: 22,
                /* 内容行高 */
                bodyrowsHeight: 22.5,
                /* 链接样式 */
                linkStyle: {},
            }
        }
```
以上参数均可根据需要配置，配置参数可参考model

## 模型
EEP 库的模型包括样式、导出字段调用格式等。

通过**EEP.model**调用

```javascript
{
            /** 样式 **/
            style: {
                /* 列字体样式 */
                font: {
                    /* 字体大小 */
                    size: 10,
                    /* 备用字体家族。整数值。 1 - Serif, 2 - Sans Serif, 3 - Mono, Others - unknown */
                    family: 1,
                    /* 字体方案 'minor', 'major', 'none' */
                    scheme: '',
                    /* 字体例如 0000FF */
                    color: { argb: '' },
                    /* 粗细体 */
                    bold: false,
                    /* 粗细 */
                    italic: '',
                    /* 删除线 */
                    strike: false,
                    /* 字体轮廓 */
                    outline: false,
                    /* 垂直对齐 */
                    vertAlign: '',
                    /* 颜色 */
                    color: { argb: '' },
                    /* 字体名 */
                    name: '',
                    /* 是否启用下划线 */
                    underline: false
                },
                /* 行充满样式 */
                fill: {
                    type: 'pattern',
                    pattern: '',
                    /* 背景色 */
                    fgColor: { argb: '' }
                },
                /* 对齐样式 */
                alignment: {
                    /* 垂直 */
                    vertical: 'top',
                    /* 水平 */
                    horizontal: 'left'
                },
                /* 边框 thin 单边框 */
                border: {
                    top: { style: '' },
                    left: { style: '' },
                    bottom: { style: '' },
                    right: { style: '' }
                },
                /* 边框可选样式
                 * thin
                 * dotted
                 * dashDot
                 * hair
                 * dashDotDot
                 * slantDashDot
                 * mediumDashed
                 * mediumDashDotDot
                 * mediumDashDot
                 * medium
                 * double
                 * thick
                 */
                borderNode: {
                    style: 'double',
                    color: { arg
                    style: 'double', color: { argb: 'FF00FF00' }
                }
            },
            /** 导出字段调用格式 **/
            field: {
                _alias: "别名/对应接口返回数据字段",
                _id: "如果涉及自定义导出字段此处需要使用同一字段",
                _name: "导出字段名",
                _show: 0,
                _width: 5,
                _vertical: 'middle',
                _horizontal: 'center'
            }
        }
```


## 工具

EEP 库还包含了一些实用的工具函数，用于处理数据和导出功能
通过**EEP.tools.XXX**调用 例如 EEP.tools.getDate
以下是当前可以调用的方法 

```javascript
			/* 获取当前时间 */
            getDate: function (type = 0) {
                // 获取当前时间的实现
            },
            /* 解析 URL */
            geturl: function (data, spstr = "||") {
                // 解析 URL 的实现
            },
            /* 处理最大信息长度 */
            splitstr: function (str) {
                // 处理最大信息长度的实现
            },
            /* 跳转到某个元素 */
            goHere: function (element, position = true) {
                // 跳转到某个元素的实现
            },
            /* 是否为允许的后缀 */
            isAllowSuffix: function (str) {
                // 检查后缀是否允许的实现
            },
            /* 生成随机串 */
            creatGuid: function getRandomCode(length = 6) {
                // 生成随机串的实现
            }
```

## 添加自定义导出字段

您可以使用以下方法向导出字段列表中添加字段：
EPP.addField("id","uid","用户ID",0,10,"middle");


```javascript
//对应以下方法
addField: function (_alias = '', _id = '', _name = '', _show = 0, _width = 5, _vertical = 'middle', _horizontal = 'center') {
            // 添加导出字段的实现
        },
```


## 导出文件

使用以下方法可以执行数据导出：
EPP.export();
```javascript
//配置
EEP.config.name="导出文件名";
//导出数据源
var content =[{id:1,name:"张三"},...{id:99,name:"李四"}]
//可根据情况选择配置的isresult=true 来要求返回本次导出数据 返回数组格式[]
var result =[];//定义为空数组
var callback;//callback为导出成功后的回调事件 不需要 不传或者传空即可
...
EPP.export(EEP.config,content,result,callback );
//对应以下方法
export: function (config = EEP.config, content = [], result = [], _callback) {
            // 数据导出的实现
        },
```
## 导出示例
以下是一个简单的示例，展示如何使用 EEP JavaScript 库进行数据导出：

```javascript
// 导出配置
        var config = {
            name: "SampleData",
            suffix: ".xlsx",
            // 其他配置项...
        };

        // 数据源
        var data = [
            // 数据项...
        ];

        // 执行导出
        EEP.export(config, data, [], function () {
            // 导出完成后的回调
            console.log("导出完成");
        });
```

## 总结
EEP JavaScript 库是一个强大的工具，用于在网页中将数据导出到 Excel 文件。通过配置选项和实用的工具函数，您可以定制导出的文件格式和样式，以满足不同需求。
由于本人为后端开发者，前端技术不能说得心应手 只能说是马马虎虎。如有问题希望各位同僚or大佬指出，谢谢
