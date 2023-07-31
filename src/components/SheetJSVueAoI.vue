<script lang="ts" setup>
import { utils, writeFile } from 'xlsx-js-style'

const header = [
    [[]], // 占位
    [
        {}, // 占位
        {
            v: `工厂统计表 ${'\n'}`,
            t: 's',
            s: {
                font: {
                    sz: 15, //设置标题的字号
                    bold: true, //设置标题是否加粗
                    name: '宋体',
                },
                //设置标题水平竖直方向居中，并自动换行展示
                alignment: {
                    horizontal: 'center',
                    vertical: 'center',
                    wrapText: true,
                },
                fill: {
                    fgColor: { rgb: '9FE3FF' },
                },
            },
        },
    ],
]

const info = [
    [
        null,
        {
            v: '                                                                                                                                                                                                                                                                                                                                                                                                    统计时间：2023/01/01 00:00',
            t: 's',
            s: {
                fill: {
                    fgColor: { rgb: '9FE3FF' },
                },
            },
        },
    ],
    [
        null,
        {
            v: '                                                                                                                                                                                                                                                                                                                                                                                                    统计维度：按月',
            t: 's',
            s: {
                fill: {
                    fgColor: { rgb: '9FE3FF' },
                },
            },
        },
    ],
    [
        null,
        {
            v: '                                                                                                                                                                                                                                                                                                                                                                                                    统计周期：2023/01/01 至 2023/01/01',
            t: 's',
            s: {
                alignment: {
                    vertical: 'top',
                },
                fill: {
                    fgColor: { rgb: '9FE3FF' },
                },
            },
        },
    ],
]

const risk = [
    '序号',
    '险别',
    '企财险',
    '家财险',
    '机动车',
    '责任险',
    '意外险',
    '货运险',
    '保证险',
    '其他险',
]

const data = risk.map((e) => {
    return {
        v: e,
        t: 's',
        s: {
            font: {
                bold: true, //设置标题是否加粗
                name: '宋体',
            },
            //设置标题水平竖直方向居中，并自动换行展示
            alignment: {
                horizontal: 'center',
                vertical: 'center',
                wrapText: true,
            },
            border: {
                top: { style: 'thin', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: '000000' } },
                right: { style: 'thin', color: { rgb: '000000' } },
            },
            fill: {
                fgColor: { rgb: '9FE3FF' },
            },
        },
    }
})

const random = (min: number, max: number): number => {
    return Math.floor(Math.random() * (max - min)) + min
}

const dataArr = () => {
    const items: (Object | null)[][] = []
    Array.apply(null, { length: 18 } as any).map((e, index: number) => {
        const item: (Object | null)[] = [null]
        Array.apply(null, { length: 10 } as any).map((ele, idx: number) => {
            item.push({
                v: idx === 0 ? index + 1 : random(1, 100000),
                t: 's',
                s: {
                    font: {
                        name: '宋体',
                    },
                    alignment: {
                        horizontal: 'center',
                        vertical: 'center',
                    },
                    border: {
                        top: { style: 'thin', color: { rgb: '000000' } },
                        bottom: { style: 'thin', color: { rgb: '000000' } },
                        left: { style: 'thin', color: { rgb: '000000' } },
                        right: { style: 'thin', color: { rgb: '000000' } },
                    },
                },
            })
        })
        items.push(item)
    })
    return items
}

const exportFile = () => {
    var ws = utils.aoa_to_sheet([
        ...header,
        ...info,
        [null, ...data],
        ...dataArr(),
    ])

    // 合并单元格
    if (!ws['!merges']) ws['!merges'] = []
    ws['!merges'].push(utils.decode_range('B2:K2'))
    ws['!merges'].push(utils.decode_range('B3:K3'))
    ws['!merges'].push(utils.decode_range('B4:K4'))
    ws['!merges'].push(utils.decode_range('B5:K5'))
    ws['!merges'].push(utils.decode_range('A2:A5'))
    ws['!merges'].push(utils.decode_range('L2:L5'))

    // 设置列宽
    // cols 为一个对象数组，依次表示每一列的宽度
    if (!ws['!cols']) ws['!cols'] = []
    ws['!cols'] = [
        { wpx: 70 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 118 },
        { wpx: 200 },
    ]

    // 设置行高
    // rows 为一个对象数组，依次表示每一行的高度
    if (!ws['!rows']) ws['!rows'] = []
    ws['!rows'] = [
        { hpx: 0 },
        { hpx: 40 },
        { hpx: 15 },
        { hpx: 15 },
        { hpx: 20 },
        { hpx: 20 },
        ...Array.apply(null, { length: dataArr().length } as any).map(() => {
            return { hpx: 20 }
        }),
    ]
    var wb = utils.book_new()
    utils.book_append_sheet(wb, ws, 'Sheet1')
    writeFile(wb, 'SheetJSVueAoO.xlsx')
}
</script>

<template>
    <div>
        <button @click="exportFile">单元格样式</button>
    </div>
</template>