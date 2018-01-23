const fs = require('fs');
const xlsx = require('node-xlsx').default;
const config = require('../config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 跨度声明，轨高变化

// 获取文件
const dir_name = __dirname.replace(/(.+\\)(.+)/, '$1')
// no.1
const beam_1 = require('./3tbom.js')
const get_file_1 = beam_1.file
const beam_data_1 = beam_1.data
// no.2
const beam_2 = require('./3tbom_1.js')
const get_file_2 = beam_2.file
const beam_data_2 = beam_2.data
// no.3
const beam_3 = require('./3tbom_2.js')
const get_file_3 = beam_3.file
const beam_data_3 = beam_3.data

// console.log(workSheetsFromBuffer[0].data[1][1])
// console.log(workSheetsFromBuffer[0].data.length)

const data = [];
// 设置表名
const sheet_name = '主梁'

// 设置首行
data.push(config.bom)

// 设置表格
for(let i = 0 ; i < beam_data_1.length ; i++){
    i > 0 ? data.push(beam_data_1[i]) : null
}
for(let i = 0 ; i < beam_data_2.length ; i++){
    i > 0 ? data.push(beam_data_2[i]) : null
}
for(let i = 0 ; i < beam_data_3.length ; i++){
    i > 0 ? data.push(beam_data_3[i]) : null
}

// 合并单元格
// const range = {s: {c: 0, r:1 }, e: {c:20, r:1}}; // A2:U2
// const range = {s: {c: 0, r:19 }, e: {c:20, r:19}}; // A20:U2
// const range = {s: {c: 0, r:37 }, e: {c:20, r:37}}; // A38:U2

const rangeArr = config.merge_cell((17-(5-0.5))/0.5*4,0,20,18)
// console.log(rangeArr)
const option = {'!merges': rangeArr}

// 输出数据
var buffer = xlsx.build([
    {
        name: sheet_name,
        data: data
    }
],option);

const t = 3
// 写入文件
fs.writeFileSync(`${dir_name}/output/${t}t` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
// fs.writeFileSync(`${dir_name}/output/`+`${file_name}` + '.'+ exportDate + '.xlsx', buffer, 'binary');
console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')
