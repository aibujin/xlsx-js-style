<script lang="ts" setup>
import { utils, writeFileXLSX } from 'xlsx-js-style'

const list = [
    { Name: 'Bill Clinton', Index: 42 },
    { Name: 'GeorgeW Bush', Index: 43 },
    { Name: 'Barack Obama', Index: 44 },
    { Name: 'Donald Trump', Index: 45 },
    { Name: 'Joseph Biden', Index: 46 },
]

const exportFile = () => {
    // 创建一个工作簿 workbook
    const workBook = utils.book_new()
    // 创建工作表 worksheet
    // json_to_sheet 	是将【由对象组成的数组】转化成sheet
    // aoa_to_sheet  	是将【一个二维数组】转化成 sheet
    // table_to_sheet  	是将【table的dom】直接转成sheet
    // 这里我们使用 json_to_sheet
    const workSheet = utils.json_to_sheet(list)
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
        <button @click="exportFile">创建工作表</button>
    </div>
</template>