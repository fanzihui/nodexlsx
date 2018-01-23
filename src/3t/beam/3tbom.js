const fs = require('fs');
const xlsx = require('node-xlsx').default;
const config = require('../config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()


// 跨度变化，轨高声明

// 获取文件
const dir_name = __dirname.replace(/(.+\\)(.+)/,'$1')
const get_file = fs.readFileSync(`${dir_name}/src/3t/BMHBOM3T.xlsx`)
// 读取文件
const workSheetsFromBuffer = xlsx.parse(get_file);

const data = [];
// 设置表名
const sheet_name = '前支腿'

// 设置首行
data.push(config.bom)

// 设置switch case 的值
const switch_case = 6
// 设置要改变数量在那一行
const switch_i = 11
// 设置要变的值中在某处跳动的值和数字
var switch_num = 8
var switch_val = 186

// 设置数量范围及基值
const switch_range = [0,5,6,7,8,9,10,11]
const switch_range_num = 3

// 设置初始值
var max = workSheetsFromBuffer[0].data.length-1,
    f_num_front = 7004,
    f_num_end = 146,
    c_num_end = 183,
    unit = '件',
    note,
    version = 00,
    span = 5,
    orbital = 4.5,
    t = 3,
    flange_num = 230

// 设置不变的值
var switch_arr = [
    '9503-00152',
    '9401-01211',
    '9401-01356',
    '9401-00825',
    '7001-00003',
    '7001-00004',
    '7004-00052',
    '7004-00187',
    '7004-00188',
    '7004-00189',
    '7004-00190',
    '7004-00191',
    '7004-00192'
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
function setxlsx(max,span,t,code,orbital,orbital_string,photo_code,is_flange) {
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
            // arr.push(workSheetsFromBuffer[0].data[i][3])
            if(is_flange){
                // switch(workSheetsFromBuffer[0].data[i][3]){
                //     case `9503-00152`:
                //     case `9401-01211`:
                //     case `9401-01356`:
                //     case `9401-00825`:
                //     case `7001-00003`:
                //     case `7001-00004`:
                //     case '7004-00052':
                //     case '7004-00187':
                //     case '7004-00188':
                //     case '7004-00189':
                //     case '7004-00190':
                //     case '7004-00191':
                //     case '7004-00192':
                //     arr.push(workSheetsFromBuffer[0].data[i][3])
                //     break;
                //     default:
                        // if(is_flange){
                        //     if( i == 17 ){
                        //         // flange_num = flange_num + 1
                        //         arr.push(`${f_num_front}-${config.five_num(flange_num)}`)
                        //     } else {
                        //         c_num_end = (c_num_end == switch_val ? c_num_end + switch_num : c_num_end + 1)
                        //         arr.push(`${f_num_front}-${config.five_num(c_num_end)}`)
                        //     }
                        // } else {
                        //     c_num_end = (c_num_end == switch_val ? c_num_end + switch_num : c_num_end + 1)
                        //     arr.push(`${f_num_front}-${config.five_num(c_num_end)}`)
                        // }
                //     break;
                // }
                // console.log(switch_arr)
                let is_switch = switch_arr.some(ele=>{
                    return ele == workSheetsFromBuffer[0].data[i][3]
                })
                if(is_switch){
                    arr.push(workSheetsFromBuffer[0].data[i][3])
                } else {
                    if( i == 17 ){
                        flange_num = flange_num + 1
                        arr.push(`${f_num_front}-${config.five_num(flange_num)}`)
                    } else {
                        c_num_end = (c_num_end == switch_val ? c_num_end + switch_num : c_num_end + 1)
                        arr.push(`${f_num_front}-${config.five_num(c_num_end)}`)
                    }
                }
            } else {
                switch(workSheetsFromBuffer[0].data[i][3]){
                    case `9503-00152`:
                    case `9401-01211`:
                    case `9401-01356`:
                    case `9401-00825`:
                    case `7001-00003`:
                    case `7001-00004`:
                    case '7004-00052':
                    case '7004-00187':
                    case '7004-00188':
                    case '7004-00189':
                    case '7004-00190':
                    case '7004-00191':
                    case '7004-00192':
                    case '7004-00193':
                        arr.push(workSheetsFromBuffer[0].data[i][3])
                    break;
                    default:
                        c_num_end = (c_num_end == switch_val ? c_num_end + switch_num : c_num_end + 1)
                        arr.push(`${f_num_front}-${config.five_num(c_num_end)}`)
                    break;
                }
            }

            // 子项名称
            arr.push(workSheetsFromBuffer[0].data[i][4])
            // 老图纸代号
            arr.push(workSheetsFromBuffer[0].data[i][5])

            // 老图纸名称
            let old_photo
            if(is_flange){
                old_photo = ( i == 17 ? '板 12X590X630' : workSheetsFromBuffer[0].data[i][6])
            } else {
                old_photo = workSheetsFromBuffer[0].data[i][6]
            }
            arr.push(old_photo)

            // 子项是否
            if( is_flange){
                // let is_child = ( i == 17 ? '是' : config.default_null)
                arr.push(config.default_null)
            } else {
                arr.push(workSheetsFromBuffer[0].data[i][7])
            }
            // 数量 变
            let product_num 
            if(photo_code == 0){
                if(i == switch_i){
                    switch(workSheetsFromBuffer[0].data[i][8]){
                        case switch_case:
                            // if(span > 4.5 && span <= 5){
                            //     product_num = 5
                            //     arr.push(5)
                            // } else if (span > 5 && span <= 6){
                            //     product_num = 6
                            //     arr.push(6)
                            // } else if (span > 6 && span <= 7){
                            //     product_num = 7
                            //     arr.push(7)
                            // } else if (span > 7 && span <= 8){
                            //     product_num = 8
                            //     arr.push(8)
                            // } else if (span > 8 && span <= 9){
                            //     product_num = 9
                            //     arr.push(9)
                            // } else if (span > 9 && span <= 10){
                            //     product_num = 10
                            //     arr.push(10)
                            // } else if (span > 10 && span <= 11){
                            //     product_num = 11
                            //     arr.push(11)
                            // } else {
                            //     product_num = workSheetsFromBuffer[0].data[i][8]
                            //     arr.push(workSheetsFromBuffer[0].data[i][8])
                            // }
                            product_num = config.range_span(switch_range,span,switch_range_num)
                            // console.log(product_num+'\n')
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
// 1张图上有2块法兰板,另加变量 flange, 若至改变法兰板参数,flange 为 true
function setExcel(span,orbital,orbital_string,photo_code,is_flange){
    let inner_span = span,
        inner_max = (17.5-(span-0.5)),
        inner_orbital = orbital
    is_flange ? c_num_end = 183 : null 
    for(let i = 0 ; i < inner_max ; i++){
        setxlsx(max,inner_span,t,f_num_end,inner_orbital,orbital_string,photo_code,is_flange);
        inner_span = (inner_span * 10 + 5) / 10
        inner_orbital = (inner_orbital * 10 + 1) / 10
        f_num_end = f_num_end + 1
    }
}

// 5˂H0≤11m
// 3t 4.5˂H0≤6m
setExcel(span,orbital,'4.4˂H0≤6m',0,false)
// 3t 6˂H0≤11m
setExcel(span,orbital,'6˂H0≤7.5m',0,true)
// 3t 11˂H0≤14m
// setExcel(span,orbital,'11˂H0≤14m',2,false)
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

