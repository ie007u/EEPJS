# EEP（Excel导出工具）

这是一个用于在浏览器中导出Excel文件的JavaScript工具。它支持多种配置选项和自定义样式，可以用于在网页上导出数据到Excel表格中。

## 配置参数

你可以通过配置以下参数来自定义导出：

```javascript
var config = {
    /* 下载文件名 */
    name: "文件名",
    /* 导出的文件后缀 */
    suffix: ".xlsx",
    /* Sheet名 */
    sheet: "Sheet",
    /* 单个excel最大承容纳数据量 */
    limit: 10000,
    /* 导出后直接下载 */
    isDown: true,
    /* 导出字段 对应自定义导出列格式 里面存放的格式为[field_format,field_format2...] */
    exportField: [],
    /* 表头占用行数 */
    headSize: 1,
    /* 冻结行数（0为不冻结） */
    freezeSize: 0,
    /* 冻结的x轴 要仅冻结行，请将其设置为 0 或 undefined */
    freezeX: 0,
    /* 冻结的y轴 要仅冻结列，请将其设置为 0 或 undefined */
    freezeY: 0,
    /* 是否启用水印 */
    isWatermark: false,
    /* 水印源 可以是base64或文件访问路径 */
    wateRmark: "",
    /* 图片格式 png jpg.. */
    wateRmarkType: "",
    /* 是否启用自动筛选 */
    isAutoScreen: false,
    /* 筛选起始位置 例如 A2 代表第一列第二行 */
    screenFrom: "",
    /* 筛选截止位置 不填则全部字段都自动筛选 */
    screenTo: "",
    /* 是否包含合并 */
    isMerge: false,
    /* 需合并的返回字段 ["1","..."] */
    mergeFiled: [],
    /* 影响合并的字 里面存放的格式为[["字段1","字段2",...],[...]] */
    mergeFiledrule: [],
    /* 是否横向延伸 */
    xExtend: false,
    /* 横向字段 */
    xExtendFiled: "",
    /* 横向前缀 请确保该前缀唯一不会与其他字段重复 */
    xExtendPrefix: "",
    /* 是否返回数据 */
    isresult: false,
    /* 临时 url样式设置 格式 A|B 或者A */
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
};
```