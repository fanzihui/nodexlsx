const path = require('path')
const fs = require('fs')
var config = {
    root: __dirname,
    no_bom: [
        '序号',
        '五位产品码',
        '规格型号',
        '产品描述',
        '起重量涵盖范围(t)',
        '跨度S范围(米)',
        '轨高H0范围(m)',
        '升高H涵盖范围(米)',
        '配置葫芦型号',
        '图纸类型',
        'BOM与MES关联',
        '新增日期',
        '备注'
    ],
    upload: [
        '物料编码',
        '规格型号',
        '工艺代码',
        '描述',
        '最小存量',
        '最大存量',
        '单位',
        '物料类型',
        '从属类别',
        '特殊属性5',
        '旧图纸名称',
        '材料',
        '旧图纸代号',
        '重量',
        '起重量涵盖范围(t)',
        '跨度S范围(米)',
        '轨高H0范围(m)',
        '升高H涵盖范围(米)',
        '配置葫芦型号',

        '备注'
    ],
    bom: [
        '序号',
        '父项代号',
        '父项名称',
        '子项代号',
        '子项名称',
        '老图纸代号',
        '老图纸名称',
        '子项是否',
        '数量',
        '单位',
        '材料',
        '单件',
        '总计',
        '备注',
        '创建日期',
        '创建人',
        '虚拟项目',
        '更改编号',
        '版本',
        '表面处理',
        '类型'
    ],
    arr_null: [
        null,
        null,
        null,
        null,
        null,
        null
    ],
    default_null: null,
    // 随机名称
    random_name() {
        let s = Math.ceil(Math.random() * 90+10)
        let now = new Date();
        let exportDate = now.getFullYear() + '.' + now.getMonth() + 1 + '.' + (now.getDate() >= 10 ? now.getDate() : '0' + now.getDate())
        return `${s}.${exportDate}`
    },
    // 合并单元格
    merge_cell(len, start, end, each) {
        var rangeArr = []
        for (let i = 0; i < len; i++) {
            let inner_r = 1,
                start_c = start,
                end_c = end
            let range = {
                s: {
                    c: start_c,
                    r: i == 0 ? 1 : each * i + inner_r
                },
                e: {
                    c: end_c,
                    r: i == 0 ? 1 : each * i + inner_r
                }
            }
            rangeArr.push(range)
        }
        return rangeArr
    },
    // 父子项代号后置码
    five_num(num) {
        let inner_num = parseInt(num)
        if (inner_num >= 100) {
            inner_num = `00${inner_num}`
        } else if (inner_num >= 10) {
            inner_num = `000${inner_num}`
        } else {
            inner_num = `0000${inner_num}`
        }
        return inner_num
    },
    // 产品编码
    product_code(code,start,mid,end,speed){
        let num = `${code}${start}-${mid > 10
            ? mid
            : '0' + mid}`
        if(end >= 10){
            num = num.concat(`-${end}`)
        } else if(end > 0){
            num = num.concat(`-0${end}`)
        } else {
            num = num.concat()
        }
        num = num.concat(speed)
        return num
    },
    // 乘 10 或 加 0
    or_ten(val) {
        return val = val >= 10 && val !== '' ? val * 10 : `0${val * 10}`
    },
    // 小数位取整,加的数与该数小数点位数一致
    deciaml(val, other = 0) {
        let count = val.toString().replace(/(\d+)\.(\d+)/, '$2')
        let sum = 1
        for (let i = 0; i <= count.length; i++) {
            sum = sum * 　10
        }
        return (val * sum + other * sum) / sum
    },
    // 小数位取整,乘的数与该数小数点位数一致
    deciaml_p(val, num) {
        let count = val.toString().replace(/(\d+)\.(\d+)/, '$2')
        let sum = 1
        for (let i = 0; i <= count.length; i++) {
            sum = sum * 　10
        }
        return (val *  num * sum) / sum
    },
    // 设置产品型号,仅针对两种葫芦
    set_mod(sheet_name, t, span, orbital, gourd, work, way) {
        // 系列名称英文 eg: BMH
        let series = sheet_name.replace(/(\w+)(.+)/, '$1')
        let inner_gourd = gourd ? 1 : 2
        let mo = `${series}${this.or_ten(t)}X${this.or_ten(span)}-${this.or_ten(orbital)}-${inner_gourd}A(D/G)-${way}-${work}`
        return mo
    },
    // 数量范围, 支腿轨高
    range(array,val,num){
        array.push(val)
        array.sort((a,b)=> a>b)
        let index = array.indexOf(val)
        array.splice(index,1)
        return 2*(num+(index-1))+2
    },
    // 数量范围, 主梁跨度
    range_span(array,val,num){
        array.push(val)
        array.sort((a,b)=> a>b)
        // console.log('调整前',array)
        let index = array.indexOf(val)
        array.splice(index,1)
        // console.log('调整后',array)
        return num+(index-1)+3
    },
    // 判断文件或者文件夹路径是否存在
    is_fileexists(path){
        fs.exists(path,(esists)=>{
            esists ? null : fs.mkdirSync(path)
        })
    }
    
}

module.exports = config