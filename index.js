#!/usr/bin/env node --max-old-space-size=8000

const yargs = require('yargs');
const xlsx = require('node-xlsx')
const fs = require('fs')
const PowerShell = require('./powershell');

let argv = yargs
    //.alias('s', 'save')
    .example('Example Url ->', 'https://www.npmjs.com/package/excelmergetool')

    //.usage('Usage: --s <filename>')
    .epilog('copyright @ linju1020@sina.com')
    .help().argv;

//合并excel表格里面的第几个工作表
var cindex = argv['i'];
if (cindex == undefined)
    cindex = 0;
console.log('i', cindex);

var cname = argv['n'];
if (cname == undefined)
    cname = "";
console.log('n', cname);

(async function () {

    var folderpath = await new PowerShell().BrowseForFolder('选择文件夹');
    console.log(`folderpath: ` + folderpath);
    folderpath += "/";

    // excel文件夹路径（把要合并的文件放在excel文件夹内）
    const _file = folderpath;//`${__dirname}/excel/`
    const _output = folderpath;//`${__dirname}/result/`
    var __name = "合并.Merge." + (cname.length > 0 ? cname : i) + ".xlsx";//'Merge'; new Date().getTime();

    // 合并数据的结果集
    let dataList = [{
        name: 'sheet1',
        data: []
    }]


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
                console.log(`开始合并：${item}`)
                if (item.indexOf('合并.Merge') == 0 || item.indexOf("~$") == 0) {
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

})();