const xlsx = require('node-xlsx')
const fs = require('fs')
const PowerShell = require('./powershell');

(async function () {

    let excelData = xlsx.parse(`./123.xlsx`);

    //console.log('excelData', );
    var data = excelData[0].data;
    console.log('data.length', data.length);


    let msg = [];

    let output = '2021-06';

    let ywTimeIndex = -1;
    let proJiesuanIndex = -1;
    let proNameIndex = -1;
    let proPriceIndex = -1;
    let proCountIndex = -1;
    data.forEach(function (item, index) {
        if (index == 0) {
            item.forEach(function (c, i) {
                if (c == '业务时间') ywTimeIndex = i;
                if (c == '结算方式') proJiesuanIndex = i;
                if (c == '商品名称') proNameIndex = i;
                if (c == '含税金额') proPriceIndex = i;
                if (c == '商品数量') proCountIndex = i;
            });
        }
        else {
            let time = item[ywTimeIndex].toString().trim(); //2021-03-08 23:59:52
            if (time.indexOf(output) == 0) {

                let jiesuan = item[proJiesuanIndex].toString().trim();
                let name = item[proNameIndex].toString().trim();
                let price = parseFloat(item[proPriceIndex]); if (isNaN(price)) price = 0;
                let count = parseFloat(item[proCountIndex]); if (isNaN(count)) count = 0;

                if (jiesuan != '货款') count = 0;

                let oldIndex = msg.findIndex(t => t.name == name);
                if (oldIndex >= 0) {
                    msg[oldIndex].price += price;
                    msg[oldIndex].count += count;
                }
                else
                    msg.push({ name, price, count });
            }
        }
    });

    console.table(msg);

    let outMsg = [];
    msg.forEach(function (item, index) {
        outMsg.push([item.name, item.price, item.count]);
    })


    // 合并数据的结果集
    let dataList = [{
        name: 'sheet1',
        data: []
    }]
    dataList[0].data = outMsg;

    //save
    // 写xlsx
    var buffer = xlsx.build(dataList)
    fs.writeFile('./' + output + '合并.xlsx', buffer, function (err) {
        if (err) {
            throw err
        }
        console.log('success');
    })

})();