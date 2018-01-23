const fs = require('fs');
const xlsx = require('node-xlsx').default;
const config = require('../config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 跨度声明，轨高变化

// 获取文件
const dir_name = __dirname.replace(/(.+\\)(.+)/, '$1')
const get_file = fs.readFileSync(`${dir_name}/src/2t/BOMLEG.xlsx`)
// 读取文件
const workSheetsFromBuffer = xlsx.parse(get_file);

// console.log(workSheetsFromBuffer[0].data[1][1])
// console.log(workSheetsFromBuffer[0].data.length)

const data = [];
// 设置表名
const sheet_name = '前支腿'

// 设置首行
data.push(config.bom)

// 设置switch case 的值
const switch_case = 10
// 设置要改变数量在那一行
const switch_i = 5
// 设置要变的值中在某处跳动的值和数字
var switch_num = 4
var switch_val = 105

// 设置数量范围及基值
const switch_range = [0,4.7,5.5,6]
const switch_range_num = 4

// 设置初始值
var max = workSheetsFromBuffer[0].data.length - 1,
    f_num_front = 7020,
    f_num_end = 2,
    c_num_end = 100,
    unit = '件',
    note,
    version = 00,
    span = 5,
    orbital = 4.5,
    t = 2

// 设置不变的值
var switch_arr = [
    `${f_num_front}-00100`, 
    `${f_num_front}-00106`, 
    `${f_num_front}-00107`, 
    `${f_num_front}-00108`
];

/**
 * 
 * @param {*循环数量} max 
 * @param {*跨度} span 
 * @param {*吨} t 
 * @param {*父项代号后置码} code 
 * @param {*轨高} orbital 
 * @param {*跨度介绍} span_string 
 * @param {*轨高介绍} orbital_string 
 * @param {*图号} photo_code 
 */
function setxlsx(max, span, t, code, orbital, span_string, orbital_string, photo_code) {
    for (let i = 0; i <= max; i++) {
        let arr = []
        if (i >= 1) {
            // 序号
            arr.push(workSheetsFromBuffer[0].data[i][0])
            // 父项代号 变
            arr.push(`${f_num_front}-${config.five_num(code)}`)

            // 父项名称
            arr.push(sheet_name)

            // 子项代号 变
            let is_switch = switch_arr.some(ele=>{
                return ele == workSheetsFromBuffer[0].data[i][3]
            })
            if(is_switch){
                arr.push(workSheetsFromBuffer[0].data[i][3])
            } else {
                c_num_end = (c_num_end == switch_val ? c_num_end + switch_num : c_num_end + 1)
                arr.push(`${f_num_front}-${config.five_num(c_num_end)}`)
            }
            // 子项名称
            arr.push(workSheetsFromBuffer[0].data[i][4])
            // 老图纸代号
            arr.push(workSheetsFromBuffer[0].data[i][5])
            // 老图纸名称
            arr.push(workSheetsFromBuffer[0].data[i][6])
            // 子项是否
            arr.push(workSheetsFromBuffer[0].data[i][7])

            // 数量 变
            let product_num
            if (i == switch_i) {
                switch (workSheetsFromBuffer[0].data[i][8]) {
                    case switch_case:
                        // if (orbital > 0 && orbital <= 4.7) {
                        //     product_num = 2 * 4 + 2
                        //     arr.push(product_num)
                        // } else if (orbital > 4.7 && orbital <= 5.5) {
                        //     product_num = 2 * 5 + 2
                        //     arr.push(product_num)
                        // } else if (orbital > 5.5 && orbital <= 6) {
                        //     product_num = 2 * 6 + 2
                        //     arr.push(product_num)
                        // }
                        product_num = config.range(switch_range,orbital,switch_range_num)
                        arr.push(product_num)
                        break;
                    default:
                        product_num = workSheetsFromBuffer[0].data[i][8]
                        arr.push(product_num)
                        break;
                }
            } else {
                product_num = workSheetsFromBuffer[0].data[i][8]
                arr.push(product_num)
            }
            // 单位
            arr.push(workSheetsFromBuffer[0].data[i][9])
            // 材料
            arr.push(workSheetsFromBuffer[0].data[i][10])
            // 单件
            arr.push(workSheetsFromBuffer[0].data[i][11])
            // 总计
            let weight = config.deciaml_p(workSheetsFromBuffer[0].data[i][11], product_num)
            arr.push(weight)
            // 备注
            arr.push(workSheetsFromBuffer[0].data[i][13])
            // 创建日期
            arr.push(workSheetsFromBuffer[0].data[i][14])
            // 创建人
            arr.push(workSheetsFromBuffer[0].data[i][15])
            // 虚拟项目
            arr.push(workSheetsFromBuffer[0].data[i][16])
            // 更改编号
            arr.push(workSheetsFromBuffer[0].data[i][17])
            // 版本
            arr.push(workSheetsFromBuffer[0].data[i][18])
            // 表面处理
            arr.push(workSheetsFromBuffer[0].data[i][19])
            // 类型
            arr.push(workSheetsFromBuffer[0].data[i][20])
        } else {
            let inner_photo_code = '1'
            let sumup = `二级BOM ${sheet_name} ${t}T，${span_string}，${((orbital * 10 - 1) / 10)}˂H0≤${orbital}（图号：Z6023${photo_code ? photo_code : inner_photo_code}）`
            arr.push(sumup)
        }
        data.push(arr)
    }
}

/**
 * 
 * @param {*跨度} span 
 * @param {*轨高} orbital 
 * @param {*跨度介绍} span_string 
 * @param {*图号} photo_code 
 */
function setExcel(span, orbital, span_string, orbital_string, photo_code) {
    let inner_span = span,
        inner_max = (6 - 4.5) * 10 + 1,
        inner_orbital = orbital
    for (let i = 0; i < inner_max; i++) {
        setxlsx(max, inner_span, t, f_num_end, inner_orbital, span_string, orbital_string, photo_code);
        inner_span = (inner_span * 10 + 5) / 10
        inner_orbital = (inner_orbital * 10 + 1) / 10
        f_num_end = f_num_end + 1
    }
}

// 2t 4.5˂S≤11m, 4.4˂H0≤6m
setExcel(span, orbital, '4.5˂S≤11m', 0)

// 合并单元格
// const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
// const range = {s: {c: 0, r:1 }, e: {c:20, r:1}}; // A2:U2
// const range = {s: {c: 0, r:11 }, e: {c:20, r:11}}; // A12:U2
// const range = {s: {c: 0, r:21 }, e: {c:20, r:21}}; // A22:U2
// const range = {s: {c: 0, r:31 }, e: {c:20, r:31}}; // A32:U2

const rangeArr = config.merge_cell((17 - (span - 0.5)) / 0.5 * 4, 0, 20, 10)
// console.log(rangeArr)
const option = {
    '!merges': rangeArr
}

// 输出数据
var buffer = xlsx.build([{
    name: sheet_name,
    data: data
}], option);

// 写入文件
if(process.env.NODE_ENV == 'dev'){
    fs.writeFileSync(`${dir_name}/output/` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
} 
console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')

// 导出相关接口
module.exports = {
    file : `${file_name}${random_name}` + '.xlsx',
    data: data,
    f_num_end: f_num_end,
    c_num_end: c_num_end
}