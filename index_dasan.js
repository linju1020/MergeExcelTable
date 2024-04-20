#!/usr/bin/env node --max-old-space-size=8000

const xlsx = require('node-xlsx')
const fs = require('fs')
// const PowerShell = require('./powershell');

var cname = ''; var cindex = 0;

// var folderpath = await new PowerShell().BrowseForFolder('选择文件夹');
// console.log(`folderpath: ` + folderpath);
// folderpath += "/";

//直接读取当前命令文件夹
var folderpath = process.cwd();
if (folderpath.substring(folderpath.length - 1) != '/') folderpath += "/";
console.log(`folderpath: ` + folderpath);

// excel文件夹路径（把要合并的文件放在excel文件夹内）
const _file = folderpath;//`${__dirname}/excel/`
const _output = folderpath;//`${__dirname}/result/`
var __name = "合并.Dasan." + (cname.length > 0 ? cname : cindex) + `.timestamp${new Date().getTime()}` + ".xlsx";//'Merge'; new Date().getTime();

// 合并数据的结果集
let dataList = [{
    name: 'sheet1',
    data: []
}]

function randomNum(minNum, maxNum) {
    switch (arguments.length) {
        case 1:
            return parseInt(Math.random() * minNum + 1, 10);
            break;
        case 2:
            return parseInt(Math.random() * (maxNum - minNum + 1) + minNum, 10);
            break;
        default:
            return 0;
            break;
    }
}

fs.readdir(_file, function (err, files) {
    if (err) {
        throw err
    }

    let totalCount = 0;
    let data_arr = [];

    // files是一个数组
    // 每个元素是此目录下的文件或文件夹的名称
    // console.log(`${files}`);
    files.forEach((item, index) => {
        try {
            //console.log(`${_file}${item}`)
            console.log(`开始打散：${item}`)
            if (item.indexOf('合并.Dasan') == 0 || item.indexOf("~$") == 0) {
                console.log('\x1B[33m%s\x1b[0m', `丢弃文件：${item}`);
                return true
            };
            if ((item.split('.')[item.split('.').length - 1]).toLowerCase() == 'xlsx' || (item.split('.')[item.split('.').length - 1]).toLowerCase() == 'xls') {
                let excelData = xlsx.parse(`${_file}${item}`)

                if (excelData) {

                    if (cname.length > 0) {
                        var Hit = false;
                        for (q = 0; q < excelData.length; q++) {
                            if (excelData[q].name == cname) {
                                cindex = q;
                                Hit = true;
                                break;
                            }
                        }
                        if (!Hit)
                            throw '没有找到 ' + cname + ' 工作表';
                    }

                    var _cData = excelData[cindex].data;
                    if (_cData.length > 0) {
                        data_arr.push(_cData);

                        console.log("length:" + _cData.length);
                        totalCount += _cData.length;
                    }

                    return true;

                }
            } else {
                console.log('\x1B[33m%s\x1b[0m', `丢弃文件：${item}`);
            }
        } catch (e) {
            console.log(e)
        }
    });

    if (data_arr.length < 1) {
        console.log(`没有任何数据`);
        return;
    }


    //console.log(data_arr);
    //return;

    //合并标题
    let CData = [];
    data_arr.forEach((item, index) => {
        item[0].forEach((tit, index2) => {
            let oldIndex = CData.findIndex(t => t == tit);
            if (oldIndex < 0) {
                CData.push(tit);
                oldIndex = CData.length - 1;
            }
            item[0][index2] = oldIndex;
        });
    });
    let _dataList = [CData];
    console.log(CData)

    //合并内容
    data_arr.forEach((item, index) => {
        item.forEach((cont_arr, index2) => {
            if (index2 == 0) return true;

            let row_data_arr = new Array(CData.length);
            for (var i = 0; i < CData.length; i++) { row_data_arr[i] = ''; }
            cont_arr.forEach((cont, index3) => {
                let oldIndex = item[0][index3];//拿到该放入的索引
                row_data_arr[oldIndex] = cont;
            });


            if (row_data_arr[9] == '货款' && row_data_arr[10] == '货款') {

                if (_dataList[0][2] == '业务主单据编码' && row_data_arr[2] != '') {
                    var _n = randomNum(11111, 99999);
                    row_data_arr[2] = row_data_arr[2].substring(0, row_data_arr[2].length - 5) + `${_n}`;
                }
                if (_dataList[0][3] == '业务子单据编码' && row_data_arr[3] != '') {
                    var _n = randomNum(1111111111, 9999999999);
                    row_data_arr[3] = row_data_arr[3].substring(0, row_data_arr[5].length - 10) + `${_n}`;
                }

                var _n2 = randomNum(111111111, 999999999);
                if (_dataList[0][4] == '交易主单' && row_data_arr[4] != '') {
                    row_data_arr[4] = row_data_arr[4].substring(0, row_data_arr[4].length - 9) + `${_n2}`;
                    row_data_arr[5] = row_data_arr[4];
                }

                if (_dataList[0][7] == '业务时间' && _dataList[0][7] != '') {

                    var minutes = randomNum(1, 180) * (randomNum(0, 1) == 0 ? -1 : 1);
                    var seconds = randomNum(1, 60) * (randomNum(0, 1) == 0 ? -1 : 1);

                    var date = new Date(Date.parse(row_data_arr[7]) + minutes * 60000 + seconds * 1000);

                    const year = date.getFullYear().toString().padStart(4, '0');
                    const month = (date.getMonth() + 1).toString().padStart(2, '0');
                    const day = date.getDate().toString().padStart(2, '0');
                    const hour = date.getHours().toString().padStart(2, '0');
                    const minute = date.getMinutes().toString().padStart(2, '0');
                    const second = date.getSeconds().toString().padStart(2, '0');
                    row_data_arr[7] = `${year}-${month}-${day} ${hour}:${minute}:${second}`
                }
            }

            _dataList.push(row_data_arr);
        });
    });

    //console.log(_dataList[0]);
    //return;
    dataList[0].data = _dataList;

    console.log('处理完毕');

    // 写xlsx
    var buffer = xlsx.build(dataList)

    let mergeFilePath = `${_output}${__name}`;


    if (fs.existsSync(mergeFilePath)) {
        //删除
        console.log('删除老文件');
        fs.unlinkSync(mergeFilePath);
    }

    console.log('开始保存');
    fs.writeFile(mergeFilePath, buffer, function (err) {
        if (err) {
            throw err
        }
        console.log('\x1B[33m%s\x1b[0m', `===================================`)
        console.log('\x1B[33m%s\x1b[0m', `完成合并：${mergeFilePath}`)

        let intCount = totalCount - data_arr.length + 1;
        let outCount = _dataList.length;

        console.log('\x1B[33m%s\x1b[0m', `导入文件数${data_arr.length}  导入行数${intCount}  合并后行数${outCount}  ${intCount == outCount ? 'Success:数据量匹配' : 'Err:数据有丢失'}`)
    })
})
