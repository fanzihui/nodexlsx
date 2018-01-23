const fs = require('fs');
const xlsx = require('node-xlsx').default;
const config = require('./config.js')
const file_name = __filename.replace(/.*\\(.*)(\.\w+)/, '$1')
const random_name = config.random_name()

// 获取文件
const get_file = fs.readFileSync(`${__dirname}/src/BOMLEG.xlsx`)
// 读取文件
const workSheetsFromBuffer = xlsx.parse(get_file);

// console.log(workSheetsFromBuffer[0].data[0]);
// console.log(workSheetsFromBuffer[0].data[1]);
// console.log(workSheetsFromBuffer[0].data[1][1])
// console.log(workSheetsFromBuffer[0].data.length)

const data = [];
// 设置表名
const sheet_name = '后支腿'

// 设置首行
data.push(config.bom)

// 设置初始值
var max = workSheetsFromBuffer[0].data.length-1,
    f_num_front = 7021,
    f_num_end = 1,
    c_num_front = 7020,
    c_num_end = 100,
    unit = '件',
    note = '借NF100-00',
    version = 00,
    span = 5,
    orbital = 4.5,
    t = 2,
    default_null = null

/**
 * 
 * @param {*循环数量} max 
 * @param {*跨度} span 
 * @param {*吨} t 
 * @param {*父项代号后置码} code 
 * @param {*轨高} orbital 
 * @param {*跨度介绍} span_string 
 * @param {*图号} photo_code 
 */
function setxlsx(max,span,t,code,orbital,span_string,photo_code) {
    for(let i = 0 ; i <= max ; i++){
        let arr = []
        if(i >= 1){
            // 序号
            arr.push(workSheetsFromBuffer[0].data[i][0])
            
            // 父项代号 变
            arr.push(`${f_num_front}-${config.five_num(code)}`)

            // 父项名称
            arr.push(sheet_name)
            
            // 子项代号 变
            // arr.push(workSheetsFromBuffer[0].data[i][3])
            switch(workSheetsFromBuffer[0].data[i][3]){
                case `${c_num_front}-00100`:
                case `${c_num_front}-00106`:
                case `${c_num_front}-00107`:
                case `${c_num_front}-00108`:
                arr.push(workSheetsFromBuffer[0].data[i][3])
                break;
                default:
                    c_num_end = (c_num_end == 105 ? c_num_end + 4 : c_num_end + 1)
                    arr.push(`${c_num_front}-${config.five_num(c_num_end)}`)
                break;
            }
            // 子项名称
            arr.push(workSheetsFromBuffer[0].data[i][4])
            // 老图纸代号
            arr.push(workSheetsFromBuffer[0].data[i][5])
            // 老图纸名称
            arr.push(workSheetsFromBuffer[0].data[i][6])
            // 子项是否
            arr.push(default_null)

            // 数量 变
            if(photo_code == 0){
                switch(workSheetsFromBuffer[0].data[i][8]){
                    case 10:
                        if(orbital > 4.7 && orbital <= 5.5){
                            arr.push(12)
                        } else if (orbital > 5.5 && orbital <= 6.4){
                            arr.push(14)
                        } else if (orbital > 6.4 && orbital <= 7.3){
                            arr.push(16)
                        } else if (orbital > 7.3){
                            arr.push(18)
                        } else {
                            arr.push(workSheetsFromBuffer[0].data[i][8])
                        }
                    break;
                    default: 
                        arr.push(workSheetsFromBuffer[0].data[i][8])
                    break;
                }
            }
            if(photo_code >= 3){
                switch(workSheetsFromBuffer[0].data[i][8]){
                    case 10:
                        if(orbital > 4.7 && orbital <= 5.6){
                            arr.push(12)
                        } else if (orbital > 5.6 && orbital <= 6.5){
                            arr.push(14)
                        } else if (orbital > 6.5 && orbital <= 7.3){
                            arr.push(16)
                        } else if (orbital > 7.3){
                            arr.push(18)
                        } else {
                            arr.push(workSheetsFromBuffer[0].data[i][8])
                        }
                    break;
                    default: 
                        arr.push(workSheetsFromBuffer[0].data[i][8])
                    break;
                }
            }

            // 单位
            arr.push(workSheetsFromBuffer[0].data[i][9])
            // 材料
            arr.push(default_null)
            // 单件
            arr.push(workSheetsFromBuffer[0].data[i][11])
            // 总计
            arr.push(workSheetsFromBuffer[0].data[i][12])
            // 备注
            arr.push(note)
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
            let inner_photo_code = orbital > 6 ? '2' : '1'
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
function setExcel(span,orbital,span_string,photo_code){
    let inner_span = span,
        inner_max = (7.5-(span-0.5))/1*10+1,
        inner_orbital = orbital
    for(let i = 0 ; i < inner_max ; i++){
        setxlsx(max,inner_span,t,f_num_end,inner_orbital,span_string,photo_code);
        inner_span = (inner_span * 10 + 5) / 10
        inner_orbital = (inner_orbital * 10 + 1) / 10
        f_num_end = f_num_end + 1
    }
}

// 2t 4.5˂S≤11米
setExcel(span,orbital,'4.5˂S≤11米',0)
// 2t 11˂S≤14米
setExcel(span,orbital,'11˂S≤14米',3)
// 2t 4.5˂S≤11米
setExcel(span,orbital,'14˂S≤17米',4)


// 合并单元格
// const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
// const range = {s: {c: 0, r:1 }, e: {c:20, r:1}}; // A2:U2
// const range = {s: {c: 0, r:11 }, e: {c:20, r:11}}; // A12:U2
// const range = {s: {c: 0, r:21 }, e: {c:20, r:21}}; // A22:U2
// const range = {s: {c: 0, r:31 }, e: {c:20, r:31}}; // A32:U2

const rangeArr = config.merge_cell((17-(span-0.5))/0.5*4,0,20,10)
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
fs.writeFileSync(`${__dirname}/output/` + `${file_name}${random_name}` + '.xlsx', buffer, 'binary');
// fs.writeFileSync(`${__dirname}/output/`+`${file_name}` + '.'+ exportDate + '.xlsx', buffer, 'binary');
console.log(`输出完毕,文件名字是: ${file_name}${random_name}` + '.xlsx')

