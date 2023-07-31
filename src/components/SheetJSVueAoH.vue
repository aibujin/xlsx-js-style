<script lang="ts" setup>
import { utils, writeFile } from 'xlsx-js-style'

const exportFile = () => {
    var ws = utils.aoa_to_sheet([
        // 特别注意合并的地方后面预留 2 个 null， 否则后面的内容（本例中是第四列其它信息）会被覆盖
        ['主要信息', null, null, '其它信息'],
        ['姓名', '性别', '年龄', '注册时间'],
        ['张三', '男', 18, new Date()],
        ['李四', '女', 22, new Date()],
    ])
    if (!ws['!merges']) ws['!merges'] = []

    // 方法一：通过 decode_range 设置范围合并单元格
    ws['!merges'].push(utils.decode_range('A1:C1'))

    // 方法二：手动设置 A1-C1 的单元格合并
    // merges 为一个对象数组，每个对象设定了单元格合并的规则
    // s:开始位置, e:结束位置, r:行, c:列

    // ws['!merges'] = [
    //     { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
    // ]

    var wb = utils.book_new()
    utils.book_append_sheet(wb, ws, 'Sheet1')
    writeFile(wb, 'SheetJSVueAoO.xlsx')
}
</script>

<template>
    <div>
        <button @click="exportFile">合并单元格</button>
    </div>
</template>