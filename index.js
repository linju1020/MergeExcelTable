#!/usr/bin/env node

const yargs = require('yargs');
const xlsx = require('node-xlsx')
const fs = require('fs')
const PowerShell = require('./powershell');

let argv = yargs
  //.alias('s', 'save')
  .example('Example Url ->', 'https://www.npmjs.com/package/MergeExcelTable')

  //.usage('Usage: --s <filename>')
  .epilog('copyright @ linju1020@sina.com')
  .help().argv;
 
(async function () {

    var folderpath = await new PowerShell().BrowseForFolder('选择文件夹');
    console.log(`folderpath: ` + folderpath);
    folderpath += "/";

    // excel文件夹路径（把要合并的文件放在excel文件夹内）
    const _file = folderpath;//`${__dirname}/excel/`
    const _output = folderpath;//`${__dirname}/result/`
    // 合并数据的结果集
    let dataList = [{
        name: 'sheet1',
        data: []
    }]


    fs.readdir(_file, function (err, files) {
        if (err) {
            throw err
        }
        // files是一个数组
        // 每个元素是此目录下的文件或文件夹的名称
        // console.log(`${files}`);
        files.forEach((item, index) => {
            try {
                // console.log(`${_file}${item}`)
                console.log(`开始合并：${item}`)
                let excelData = xlsx.parse(`${_file}${item}`)
                if (excelData) {

                    //
                    console.log("length:" + excelData[0].data.length)
                    /* for (var c in excelData[0].data) {
                        var comname = excelData[0].data[c][10]
                            //console.log(comname);
                        if (comname != '***有限公司-寄售' && comname != '供应商名称') {
                            console.error(item + "err:" + comname);
                            return;
                        }
                    }  */ 
                    //

                    if (dataList[0].data.length > 0) {
                        excelData[0].data.splice(0, 1)
                    }
                    dataList[0].data = dataList[0].data.concat(excelData[0].data)
                }
            } catch (e) {
                console.log(e)
                console.log('excel表格内部字段不一致，请检查后再合并。')
            }
        })
        // 写xlsx
        var buffer = xlsx.build(dataList)
        var __name = new Date().getTime();
        fs.writeFile(`${_output}合并.${__name}.xlsx`, buffer, function (err) {
            if (err) {
                throw err
            }
            console.log('\x1B[33m%s\x1b[0m', `完成合并：${_output}合并.${__name}.xlsx`)
        })
    })

})();