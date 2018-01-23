const fs = require('fs');
const xlsx = require('node-xlsx').default;
const dir_name = __dirname.replace(/(.+\\)(.+)/, '$1')
const config = require(`${dir_name.replace(/(.+\\)(src.+)/ig,'$1')}config.js`)
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 跨度变化，轨高声明

// 获取文件
const get_file = fs.readFileSync(`${dir_name}/assets/BEAM_1.xlsx`)

// 读取文件
const workSheetsFromBuffer = xlsx.parse(get_file);

const data = [];
// 设置表名
const sheet_name = '主梁'

// 设置首行
data.push(config.bom)

// 设置switch case 的值
const switch_case = 13
// 设置要改变数量在那一行
const switch_i = 11
// 设置要变的值中在某处跳动的值和数字
var switch_num = 8
var switch_val = 245

// 设置数量范围及基值
const switch_range = [0,12,13,14]
const switch_range_num = 10

// 引入上一个文件,获取前置码,后置码
const leg_data  = require('./beam_0')

// console.log(leg_data.c_num_end)
// 设置初始值
var max = workSheetsFromBuffer[0].data.length-1,
    f_num_front = leg_data.f_num_front,
    f_num_end = leg_data.f_num_end,
    c_num_end = leg_data.c_num_end,
    unit = '件',
    note,
    version = 00,
    span = 11.5,
    orbital = 4.5,
    t = 3

// 设置不变的值
var switch_arr = [
    '9503-00152',
    '9401-01211',
    '9401-01356',
    '9401-00825',
    '7001-00003',
    '7001-00004',
    '7004-00052',
    '7004-00246',
    '7004-00247',
    '7004-00248',
    '7004-00249',
    '7004-00250',
    '7004-00251',
    '7004-00252',
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
function setxlsx(max,span,t,code,orbital,orbital_string,photo_code) {
    for(let i = 0 ; i <= max ; i++){
        let arr = []
        if(i >= 1){
            // 序号
            arr.push(workSheetsFromBuffer[0].data[i][0])
            // 父项代号 变 前置码-后置码
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
            arr.push( workSheetsFromBuffer[0].data[i][6])

            // 子项是否
            arr.push(workSheetsFromBuffer[0].data[i][7])
            // 数量 变
            let product_num 
            if(i == switch_i){
                switch(workSheetsFromBuffer[0].data[i][8]){
                    case switch_case:
                        product_num = config.range_span(switch_range,span,switch_range_num)
                        arr.push(product_num)
                    break;
                    default: 
                        product_num = workSheetsFromBuffer[0].data[i][8]
                        arr.push(workSheetsFromBuffer[0].data[i][8])
                    break;
                }
            } else {
                product_num = workSheetsFromBuffer[0].data[i][8]
                arr.push(workSheetsFromBuffer[0].data[i][8])
            }
            // 单位
            arr.push(workSheetsFromBuffer[0].data[i][9])
            // 材料
            arr.push(workSheetsFromBuffer[0].data[i][10])
            // 单件
            arr.push(workSheetsFromBuffer[0].data[i][11])
            // 总计 单件 x 数量
            let weight = config.deciaml_p(workSheetsFromBuffer[0].data[i][11],product_num)
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
            let inner_photo_code = photo_code == 0 ? 1 : 2;
            let sumup = `二级BOM ${sheet_name} ${t}T，${((span * 10 - 5) / 10)}˂S≤${span}，${orbital_string}（图号：M60${t}3${photo_code ? photo_code : inner_photo_code}）`
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
 * @param {*轨高介绍} orbital_string 
 * @param {*图号} photo_code 
 */
function setExcel(span,orbital,orbital_string,photo_code){
    let inner_span = span,
        inner_max = (17-(span-0.5)),
        inner_orbital = orbital
    for(let i = 0 ; i < inner_max ; i++){
        setxlsx(max,inner_span,t,f_num_end,inner_orbital,orbital_string,photo_code);
        inner_span = (inner_span * 10 + 5) / 10
        inner_orbital = (inner_orbital * 10 + 1) / 10
        f_num_end = f_num_end + 1
    }
}


// 3t 11˂H0≤14m
setExcel(span,orbital,'4.4˂H0≤7.5m',2)
// 3t 14˂H0≤17m
// setExcel(span,orbital,'6˂H0≤7.5m',0,true)

// 合并单元格
// const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
// const range = {s: {c: 0, r:1 }, e: {c:20, r:1}}; // A2:U2
// const range = {s: {c: 0, r:11 }, e: {c:20, r:11}}; // A12:U2
// const range = {s: {c: 0, r:21 }, e: {c:20, r:21}}; // A22:U2
// const range = {s: {c: 0, r:31 }, e: {c:20, r:31}}; // A32:U2

// const rangeArr = config.merge_cell((17-(span-0.5))/0.5*4,0,20,10)
const rangeArr = {}
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
// const output = dir_name.replace(/(.+\\)(src.+)/ig,'$1')
// 手动修改是否联动
var global_test = false

if(process.env.NODE_ENV == 'dev' && global_test){
    fs.writeFileSync(`${config.root}/output/${t}t/` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
    console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')
} else{
    console.log(`检测完毕,可以输出: ${file_name}${random_name}` + '.xlsx')
}

// 导出相关接口
module.exports = {
    file : `${file_name}${random_name}` + '.xlsx',
    data: data,
    f_num_end: f_num_end,
    c_num_end: c_num_end,
    f_num_front: f_num_front,
}

