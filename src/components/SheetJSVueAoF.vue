<script lang="ts" setup>
import { utils, writeFileXLSX } from 'xlsx-js-style'

const list = [{ Name: 'Vue' }, { Name: 'React' }, { Name: 'Angular' }]

const exportFile = () => {
    // 创建一个工作簿 workbook
    const workBook = utils.book_new()
    // 创建工作表 worksheet
    const workSheet = utils.json_to_sheet(list)

    // 添加注释
    if (!workSheet.A2.c) workSheet.A2.c = []
    workSheet.A2.c.push({
        a: 'Vue',
        t: 'Vue 是一款用于构建用户界面的 JavaScript 框架',
    })

    if (!workSheet.A3.c) workSheet.A3.c = []
    // 如果设置为 true，则只有当用户将鼠标悬停在注释上时，注释才可见；
    workSheet.A3.c.hidden = true
    workSheet.A3.c.push({
        a: 'React',
        t: 'React 用于构建 Web 和原生交互界面的库',
    })

    if (!workSheet.A4.c) workSheet.A4.c = []
    // 如果设置为 true，则只有当用户将鼠标悬停在注释上时，注释才可见；
    workSheet.A4.c.hidden = true
    workSheet.A4.c.push({
        a: 'Angular',
        t: 'Angular 是一个应用设计框架与开发平台，旨在创建高效而精致的单页面应用',
    })

    // 将工作表放入工作簿中
    utils.book_append_sheet(workBook, workSheet, 'Data')

    // 生成数据保存
    writeFileXLSX(workBook, `SheetJSVueAoO.xlsx`, {
        bookType: 'xlsx',
    })
}
</script>

<template>
    <div>
        <button @click="exportFile">单元格注释</button>
    </div>
</template>