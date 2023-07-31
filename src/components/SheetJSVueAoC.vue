<script lang="ts" setup>
import { utils, writeFileXLSX } from 'xlsx-js-style'

const list = [
    { Name: 'Barack Obama', Taxes: 726223 },
    { Name: 'GeorgeW Bush', Taxes: 3.5 },
    { Name: 'Bill Clinton', Taxes: 45571 },
    { Name: 'Donald Trump', Taxes: 0.0219 },
    { Name: 'Donald Trump', Taxes: new Date() },
    { Name: 'Joseph Biden', Taxes: 666666 },
]

const exportFile = () => {
    // 创建一个工作簿 workbook
    const workBook = utils.book_new()
    // 创建工作表 worksheet
    const workSheet = utils.json_to_sheet(list)

    // 分配数字格式
    workSheet['B3'].z = '"$"#,##0.00_);\\("$"#,##0.00\\)'
    workSheet['B4'].z = '#,##0'
    workSheet['B5'].z = '0.00%'
    workSheet['B6'].z = 'yyyy-mm-dd hh:mm:ss AM/PM'
    workSheet['B7'].z = '[Red](#,##0)'

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
        <button @click="exportFile">分配数字格式</button>
    </div>
</template>