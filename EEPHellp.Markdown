# EEP��Excel�������ߣ�

����һ��������������е���Excel�ļ���JavaScript���ߡ���֧�ֶ�������ѡ����Զ�����ʽ��������������ҳ�ϵ������ݵ�Excel�����С�

## ���ò���

�����ͨ���������²������Զ��嵼����

```javascript
var config = {
    /* �����ļ��� */
    name: "�ļ���",
    /* �������ļ���׺ */
    suffix: ".xlsx",
    /* Sheet�� */
    sheet: "Sheet",
    /* ����excel�������������� */
    limit: 10000,
    /* ������ֱ������ */
    isDown: true,
    /* �����ֶ� ��Ӧ�Զ��嵼���и�ʽ �����ŵĸ�ʽΪ[field_format,field_format2...] */
    exportField: [],
    /* ��ͷռ������ */
    headSize: 1,
    /* ����������0Ϊ�����ᣩ */
    freezeSize: 0,
    /* �����x�� Ҫ�������У��뽫������Ϊ 0 �� undefined */
    freezeX: 0,
    /* �����y�� Ҫ�������У��뽫������Ϊ 0 �� undefined */
    freezeY: 0,
    /* �Ƿ�����ˮӡ */
    isWatermark: false,
    /* ˮӡԴ ������base64���ļ�����·�� */
    wateRmark: "",
    /* ͼƬ��ʽ png jpg.. */
    wateRmarkType: "",
    /* �Ƿ������Զ�ɸѡ */
    isAutoScreen: false,
    /* ɸѡ��ʼλ�� ���� A2 ������һ�еڶ��� */
    screenFrom: "",
    /* ɸѡ��ֹλ�� ������ȫ���ֶζ��Զ�ɸѡ */
    screenTo: "",
    /* �Ƿ�����ϲ� */
    isMerge: false,
    /* ��ϲ��ķ����ֶ� ["1","..."] */
    mergeFiled: [],
    /* Ӱ��ϲ����� �����ŵĸ�ʽΪ[["�ֶ�1","�ֶ�2",...],[...]] */
    mergeFiledrule: [],
    /* �Ƿ�������� */
    xExtend: false,
    /* �����ֶ� */
    xExtendFiled: "",
    /* ����ǰ׺ ��ȷ����ǰ׺Ψһ�����������ֶ��ظ� */
    xExtendPrefix: "",
    /* �Ƿ񷵻����� */
    isresult: false,
    /* ��ʱ url��ʽ���� ��ʽ A|B ����A */
    urltag: "",
    /* �������� */
    /* Ĭ��ȫ����ʽ */
    defaultStyle: {
        /* ͷ������ʽ */
        headcolfont: {},
        /* ͷ������ʽ */
        headrowfill: {},
        /* ����ʽ */
        bodycolfont: {},
        /* ��������ʽ */
        bodyrowfill: {},
        /* ͷ���и� */
        headrowsHeight: 22,
        /* �����и� */
        bodyrowsHeight: 22.5,
        /* ������ʽ */
        linkStyle: {},
    }
};
```