const fs = require('fs');
const xlsx = require('node-xlsx').default;
const config = require('../config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 跨度声明，轨高变化

// 获取文件
const dir_name = __dirname.replace(/(.+\\)(.+)/, '$1')
// no.1
const leg_0 = require('./bom_leg.js')
const get_file_0 = leg_0.file
const leg_data_0 = leg_0.data
// no.2
const leg_1 = require('./2tbom_leg_1.js')
const get_file_1 = leg_1.file
const leg_data_1 = leg_1.data
// no.3
const leg_2 = require('./2tbom_leg_2.js')
const get_file_2 = leg_2.file
const leg_data_2 = leg_2.data
// no.4
const leg_3 = require('./2tbom_leg_3.js')
const get_file_3 = leg_3.file
const leg_data_3 = leg_3.data

// console.log(workSheetsFromBuffer[0].data[1][1])
// console.log(workSheetsFromBuffer[0].data.length)

const data = [];
// 设置表名
const sheet_name = '前支腿'

// 设置首行
data.push(config.bom)

// 设置表格
for(let i = 0 ; i < leg_data_0.length ; i++){
    i > 0 ? data.push(leg_data_0[i]) : null
}
for(let i = 0 ; i < leg_data_1.length ; i++){
    i > 0 ? data.push(leg_data_1[i]) : null
}
for(let i = 0 ; i < leg_data_2.length ; i++){
    i > 0 ? data.push(leg_data_2[i]) : null
}
for(let i = 0 ; i < leg_data_3.length ; i++){
    i > 0 ? data.push(leg_data_3[i]) : null
}

// 合并单元格
// const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
// const range = {s: {c: 0, r:1 }, e: {c:20, r:1}}; // A2:U2
// const range = {s: {c: 0, r:11 }, e: {c:20, r:11}}; // A12:U2
// const range = {s: {c: 0, r:21 }, e: {c:20, r:21}}; // A22:U2
// const range = {s: {c: 0, r:31 }, e: {c:20, r:31}}; // A32:U2

const rangeArr = config.merge_cell((17-(5-0.5))/0.5*4,0,20,10)
// console.log(rangeArr)
const option = {'!merges': rangeArr}

// 输出数据
var buffer = xlsx.build([
    {
        name: sheet_name,
        data: data
    }
],option);

// 写入文件
fs.writeFileSync(`${dir_name}/output/2t` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
// fs.writeFileSync(`${dir_name}/output/`+`${file_name}` + '.'+ exportDate + '.xlsx', buffer, 'binary');
console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')
