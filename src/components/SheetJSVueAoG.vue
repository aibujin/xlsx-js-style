<script lang="ts" setup>
import { utils, writeFileXLSX, writeFile } from 'xlsx-js-style'

const list = [
    ['姓名', '语文', '数学', '英语', '总数', '最大值', '姓名去重'],
    ['张三', 80, 100, 100],
    ['李四', 90, 100, 80],
    ['李四', 85, 80, 100],
    ['王五', 100, 85, 90],
    ['张三', 90, 70, 90],
    ['赵六', 95, 90, 80],
    ['张三', 100, 80, 90],
]

const exportSimpleFormula = () => {
    var ws = utils.aoa_to_sheet([
        [6], // A1
        [8], // A2
        [{ t: 'n', v: 3, f: 'SUM(A1,A2)' }], // SUM 函数
        [{ t: 'n', v: 3, f: 'CONCAT("concat：",A1,A2)' }], // CONCAT 函数
    ])
    var wb = utils.book_new()
    utils.book_append_sheet(wb, ws, 'Sheet1')
    writeFile(wb, 'SheetJSFormula.xlsx')
}

const exportFile = () => {
    // 创建一个工作簿 workbook
    const workBook = utils.book_new()
    // 创建工作表 worksheet
    const workSheet = utils.aoa_to_sheet(list)

    utils.sheet_set_array_formula(workSheet, 'E2:E8', 'B2:B8+C2:C8+D2:D8', true)

    list.forEach((e: (string | number)[], index: number) => {
        if (index > 0) {
            utils.sheet_set_array_formula(
                workSheet,
                `F${index + 1}`,
                `MAX(B${index + 1},C${index + 1},D${index + 1})`,
                true
            )
        }
    })

    utils.sheet_set_array_formula(
        workSheet,
        'G2:G8',
        '_xlfn.UNIQUE(A2:A8)',
        true
    )

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
        <button @click="exportSimpleFormula">简单公式</button>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <button @click="exportFile">数组公式</button>
    </div>
</template>