<script lang="ts" setup>
import { utils, writeFileXLSX } from 'xlsx-js-style'

const list = [
    {
        Name: 'Bill Clinton',
        Date: '2023-01-01',
        'Source category name': 'Excise Taxes',
        'Source subcategory name': 'Corporation Income Taxes',
    },
    {
        Name: 'GeorgeW Bush',
        Date: '2023-01-01',
        'Source category name': 'Excise Taxes',
        'Source subcategory name': 'Corporation Income Taxes',
    },
    {
        Name: 'Barack Obama',
        Date: '2023-01-01',
        'Source category name': 'Excise Taxes',
        'Source subcategory name': 'Corporation Income Taxes',
    },
    {
        Name: 'Donald Trump',
        Date: '2023-01-01',
        'Source category name': 'Excise Taxes',
        'Source subcategory name': 'Corporation Income Taxes',
    },
    {
        Name: 'Joseph Biden',
        Date: '2023-01-01',
        'Source category name': 'Excise Taxes',
        'Source subcategory name': 'Corporation Income Taxes',
    },
]

const exportFile = () => {
    // 创建一个工作簿 workbook
    const workBook = utils.book_new()
    // 创建工作表 worksheet
    const workSheet = utils.json_to_sheet(list)

    // 设置列宽
    // cols 为一个对象数组，依次表示每一列的宽度
    // wpx 字段表示以像素为单位，wch 字段表示以字符为单位
    // hidden 如果为真，则隐藏该列
    workSheet['!cols'] = [
        { wpx: 100 },
        { wch: 50 },
        { width: 30 },
        { hidden: true },
    ]

    // 设置行高
    // rows 为一个对象数组，依次表示每一行的高度
    workSheet['!rows'] = [{ hpx: 30 }, { hpt: 50 }, { hidden: true }]

    // 将工作表放入工作簿中
    // utils.book_append_sheet(workbook, worksheet, name, true);
    utils.book_append_sheet(workBook, workSheet, 'Data')
    // 生成数据保存
    writeFileXLSX(workBook, `SheetJSVueAoO.xlsx`, {
        bookType: 'xlsx',
    })
}
</script>

<template>
    <div>
        <button @click="exportFile">设置单元格宽度/高度/隐藏</button>
    </div>
</template>