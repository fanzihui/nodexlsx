const fs = require('fs');
const xlsx = require('node-xlsx').default;
const dir_name = __dirname.replace(/(.+\\)(.+)/, '$1')
const config = require(`${dir_name.replace(/(.+\\)(src.+)/ig,'$1')}config.js`)
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 跨度变化，轨高声明

// 获取文件
const get_file = fs.readFileSync(`${dir_name}/assets/BEAM.xlsx`)
// 读取文件
const workSheetsFromBuffer = xlsx.parse(get_file);

const data = [];
// 设置表名
const sheet_name = '主梁'

// 设置首行
data.push(config.bom)

// 引入上一个文件,获取前置码,后置码
const beam_data = require(`${config.root}/src/10t/beam/beam_2.js`)
console.log('f_num_end', beam_data.f_num_end)
console.log('c_num_end', beam_data.c_num_end)

// 设置switch case 的值
const switch_case = 6
// 设置要改变数量在那一行
const switch_i = 11
// 设置要变的值中在某处跳动的值和数字
var switch_num = 8
// 321是在此处要跳动的值
var switch_val = 480

// 设置数量范围及基值
const switch_range = [0, 5, 6, 7, 8, 9, 10, 11]
const switch_range_num = 3

// 设置初始值
var max = workSheetsFromBuffer[0].data.length - 1,
    f_num_front = 7004,
    f_num_end = beam_data.c_num_end + 1,
    c_num_end = 477,
    unit = '件',
    note,
    version = 00,
    span = 5,
    orbital = 4.5,
    t = 10,
    // 第一次循环完的最后一个值 
    flange_num = 523

// 设置不变的值
var switch_arr = [
    '9503-00152',
    '9401-01211',
    '9401-01356',
    '9401-00825',
    '7001-00003',
    '7001-00004',
    '7004-00052',
    '7004-00481',
    '7004-00482',
    '7004-00483',
    '7004-00484',
    '7004-00485',
    '7004-00486',
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
function setxlsx(max, span, t, code, orbital, orbital_string, photo_code, is_flange) {
    for (let i = 0; i <= max; i++) {
        let arr = []
        if (i >= 1) {
            // 序号
            arr.push(workSheetsFromBuffer[0].data[i][0])
            // 父项代号 变 前置码-后置码
            arr.push(`${f_num_front}-${config.five_num(code)}`)
            // 父项名称
            arr.push(sheet_name)
            // 子项代号 变
            if (is_flange) {
                let is_switch = switch_arr.some(ele => {
                    return ele == workSheetsFromBuffer[0].data[i][3]
                })
                if (is_switch) {
                    arr.push(workSheetsFromBuffer[0].data[i][3])
                } else {
                    if (i == 17) {
                        flange_num = flange_num + 1
                        arr.push(`${f_num_front}-${config.five_num(flange_num)}`)
                    } else {
                        c_num_end = (c_num_end == switch_val ? c_num_end + switch_num : c_num_end + 1)
                        arr.push(`${f_num_front}-${config.five_num(c_num_end)}`)
                    }
                }
            } else {
                switch (workSheetsFromBuffer[0].data[i][3]) {
                    case `9503-00152`:
                    case `9401-01211`:
                    case `9401-01356`:
                    case `9401-00825`:
                    case `7001-00003`:
                    case `7001-00004`:
                    case '7004-00052':
                    case '7004-00481':
                    case '7004-00482':
                    case '7004-00483':
                    case '7004-00484':
                    case '7004-00485':
                    case '7004-00486':
                    case '7004-00487':
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
            if (is_flange) {
                old_photo = (i == 17 ? '板 12×640×950' : workSheetsFromBuffer[0].data[i][6])
            } else {
                old_photo = workSheetsFromBuffer[0].data[i][6]
            }
            arr.push(old_photo)

            // 子项是否
            if (is_flange) {
                // let is_child = ( i == 17 ? '是' : config.default_null)
                arr.push(config.default_null)
            } else {
                arr.push(workSheetsFromBuffer[0].data[i][7])
            }
            // 数量 变
            let product_num
            if (photo_code == 0) {
                if (i == switch_i) {
                    switch (workSheetsFromBuffer[0].data[i][8]) {
                        case switch_case:
                            product_num = config.range_span(switch_range, span, switch_range_num)
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
            let inner_photo_code = photo_code == 0 ? 1 : 2;
            let sumup = `二级BOM ${sheet_name} ${t}T，${((span * 10 - 5) / 10)}˂S≤${span}，${orbital_string}（图号：M6${t >= 10 ? t : '0' + t}3${photo_code ? photo_code : inner_photo_code}）`
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
function setExcel(span, orbital, orbital_string, photo_code, is_flange) {
    let inner_span = span,
        inner_max = (17.5 - (span - 0.5)),
        inner_orbital = orbital
    is_flange ? c_num_end = 477 : null
    for (let i = 0; i < inner_max; i++) {
        setxlsx(max, inner_span, t, f_num_end, inner_orbital, orbital_string, photo_code, is_flange);
        inner_span = (inner_span * 10 + 5) / 10
        inner_orbital = (inner_orbital * 10 + 1) / 10
        f_num_end = f_num_end + 1
    }
}

// 5˂H0≤11m
// 3t 4.5˂H0≤6m
setExcel(span, orbital, '4.4˂H0≤6m', 0, false)
// 3t 6˂H0≤11m
setExcel(span, orbital, '6˂H0≤7.5m', 0, true)

// const rangeArr = config.merge_cell((17-(span-0.5))/0.5*4,0,20,10)
const rangeArr = []
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
// const output = dir_name.replace(/(.+\\)(src.+)/ig,'$1')
// 手动修改是否联动
var global_test = 0

if (process.env.NODE_ENV == 'dev' && global_test) {
    fs.writeFileSync(`${config.root}/output/${t}t/` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
    console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')
} else {
    console.log(`检测完毕,可以输出: ${file_name}${random_name}` + '.xlsx')
}

// 导出相关接口
module.exports = {
    file: `${file_name}${random_name}` + '.xlsx',
    data: data,
    f_num_end: f_num_end,
    c_num_end: flange_num,
    f_num_front: f_num_front,
    t: t,
}