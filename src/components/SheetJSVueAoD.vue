<script lang="ts" setup>
import { utils, writeFileXLSX } from 'xlsx-js-style'

const list = [
    { Name: 'https://sheetjs.com' },
    { Name: '电子邮箱' },
    { Name: '访问 C 盘文件' },
    { Name: '选中指定单元格' },
    { Name: '跳转指定 Sheet' },
    { Name: 'Joseph Biden' },
]

const exportFile = () => {
    // 创建一个工作簿 workbook
    const workBook = utils.book_new()
    // 创建工作表 worksheet
    const workSheet = utils.json_to_sheet(list)

    // 链接 https://sheetjs.com
    workSheet['A2'].l = {
        Target: 'https://sheetjs.com',
        Tooltip: 'https://sheetjs.com',
    }

    // 链接电子邮箱
    workSheet['A3'].l = { Target: 'mailto:ignored@dev.null' }

    // 访问本地 C 盘文件
    workSheet['A4'].l = { Target: 'file:///C:/Users/pc/Downloads/receipts.xls' }

    // 选中指定单元格 A1:C5
    workSheet['A5'].l = { Target: '#A1:C5', Tooltip: '选中 A1:C5 ' }

    // 跳转指定 Sheet
    workSheet['A6'].l = { Target: '#Data2!A1:C6', Tooltip: 'Data2' }

    workSheet['A7'].l = { Target: '#SheetJSDN', Tooltip: 'Defined Name' }

    // 将工作表放入工作簿中
    // utils.book_append_sheet(workbook, worksheet, name, true);
    utils.book_append_sheet(workBook, workSheet, 'Data')

    // 创建工作表2 worksheet
    var worksheet2 = utils.aoa_to_sheet([['Same', 'Cross', 'Name']])
    utils.book_append_sheet(workBook, worksheet2, 'Data2')

    // 定义的名称, ref 选中的是当前超链接所在单元格位置
    workBook.Workbook = {
        Names: [{ Name: 'SheetJSDN', Ref: 'Data2!A1:A1' }],
    }

    // 生成数据保存
    writeFileXLSX(workBook, `SheetJSVueAoO.xlsx`, {
        bookType: 'xlsx',
    })
}
</script>

<template>
    <div>
        <button @click="exportFile">超链接</button>
    </div>
</template>