const xlsx = require('node-xlsx').default
const fs = require('fs')
const config = require('./config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

const data = [];
data.push(config.no_bom)

const sheet_name = 'BMH系列'
const arrNull = config.arr_null;
/**
 * 设置初始值
 * id 序列号
 * max 循环次数
 * span 跨度
 * mod 产品模型中间数字,跨度变,吨位未变
 * work 工作级别
 * way 操作方式
 * photo 图纸类型
 * bomMes BOM与MES关联
 * note 备注
 */

var defaultMo = true,
    id = 0,
    max = 31,
    t,
    span = 5,
    mod = 0,
    work = 'A4',
    way = '地操',
    photo = '标准图纸',
    bomMes = '导入系统',
    note = 'D/G(20/30)m/min',
    date = '2018.01.08',
    speed_code = '-D/G'

/**
 * 
 * @param {*} defaultMo 
 * @param {*} num 
 * @param {*} t 
 * @param {*} max 
 * @param {*} id 
 * @param {*} span 
 * @param {*} mod 
 */
function setXlsx(defaultMo, num, t, max, id, span, mod) {

    /**
     * inner_id
     * models
     * cd1
     * inner_max
     * orbital 轨高
     * height 升高涵盖范围
     * t 吨位
     *
     */
    var inner_id = id,
        models = mod,
        inner_span = span,
        cd1 = defaultMo,
        inner_max = max,
        inner_t = t,
        inner_num = num,
        orbital = 4.5,
        heigh = 6

    for (let i = 0; i < inner_max; i++) {
        let arr = []
        // 序号
        arr.push(i + 1);

        // 五位产品码
        arr.push(config.product_code('NF',inner_num,models,inner_id,speed_code))
        // 产品型号
        arr.push(config.set_mod(sheet_name,t,span,orbital,cd1,work,'1F'))

        let set_heigh
        switch (inner_t) {
            case 2:
                set_heigh = orbital > 7 ? `6＜H≤9` : `0＜H≤6`
                heigh = orbital > 7 ? 9 : 6
                break;
            case 3:
                set_heigh = orbital > 7.2 ? `6＜H≤9` : `0＜H≤6`
                heigh = orbital > 7.2 ? 9 : 6
                break;
            case 5:
                set_heigh = orbital > 7.4 ? `6＜H≤9` : `0＜H≤6`
                heigh = orbital > 7.4 ? 9 : 6
                break;
            case 10:
            case 16:
                set_heigh = `0＜H≤9`
                heigh = 9
                break;
            default:
                break;
        }
        // 介绍
        let desc = `起重量${inner_t}t,跨度${span}m,轨高${orbital}米,葫芦型号${cd1
            ? 'CD1'
            : 'MD1'}型${inner_t}t${heigh}m,${way}（分低速/高速）,工作级别${work}。
        `
        arr.push(desc)

        // 起重量涵盖范围
        switch (inner_t) {
            case 2:
                arr.push('t≤' + inner_t)
                break;
            case 3:
                arr.push(`2＜t≤${inner_t}`)
                break;
            case 5:
                arr.push(`3＜t≤${inner_t}`)
                break;
            case 10:
                arr.push(`5＜t≤${inner_t}`)
                break;
            case 16:
                arr.push(`10＜t≤${inner_t}`)
                break;
            default:
                break;
        }

        // 跨度范围
        let set_span = (span - 0.5) + '＜S≤' + span
        arr.push(set_span)
        // 轨高范围
        let set_orbital = ((orbital * 10 - 1) / 10) + '<H0≤' + (orbital)
        arr.push(set_orbital)

        // 升高
        arr.push(set_heigh)
        // 配型葫芦
        let set_cdmd = `${cd1 ? 'CD1' : 'MD1'}型${t}t-${heigh}m`

        arr.push(set_cdmd)
        arr.push(photo)
        arr.push(bomMes)
        arr.push(date)
        arr.push(note)
        // 循环变量
        orbital = (orbital * 10 + 1) / 10;
        inner_id++;
        data.push(arr);
    }
}

function set_excel(t, num, max, id, span) {
    let inner_span = span
    for (let j = 0; j < 25; j++) {
        setXlsx(1, num, t, max, id, inner_span, j)
        setXlsx(0, num, t, max, 31, inner_span, j)
        inner_span = (inner_span * 10 + 5) / 10
        data.push(arrNull);
    }
}

// 设置数据
// 2t
set_excel(2, 100, max, id, span)
// 3t
set_excel(3, 101, max, id, span)
// 5t
set_excel(5, 102, max, id, span)
// 10t
set_excel(10, 103, max, id, span)
// 16t
set_excel(16, 104, max, id, span)

// 输出数据
var buffer = xlsx.build([{
    name: sheet_name,
    data: data
}]);

fs.writeFileSync(`${__dirname}/output/` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
// fs.writeFileSync(`${__dirname}/output/`+`${file_name}` + '.'+ exportDate + '.xlsx', buffer, 'binary');
console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')