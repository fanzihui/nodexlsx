const fs = require('fs');
const xlsx = require('node-xlsx').default;
const config = require('../../../config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 跨度声明，轨高变化

const data = [];
// 设置表名
const sheet_name = '前支腿'

// 设置首行
data.push(config.bom)

// 扫描文件
const cur_dir = __dirname
var file_arr = fs.readdirSync(cur_dir, (err, data) => {
    if (err) throw err;
    return data
});

file_arr.forEach((ele,index) => {
    ele.search(/merge/ig) != -1 ? file_arr.splice(index,1): null;
});
// 重新排序
file_arr.sort()

// 设置表格
for(let i = 0 ; i < file_arr.length ;i++){
    let filename = `./${file_arr[i]}`
    // console.log(filename)
    var leg = require(filename)
    var leg_data = leg.data
    for(let j = 0; j < leg_data.length; j++){
        j > 0 ? data.push(leg_data[j]) : null
    }
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

const t = 3
// 写入文件
config.is_fileexists(`${config.root}/output/${t}t/`)
fs.writeFileSync(`${config.root}/output/${t}t/${file_name}${random_name}` + '.xlsx', buffer, 'binary');
console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')