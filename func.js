var http = require("http");

function download(url, callback) {
    http.get(url, function (res) {
        var data = "";
        res.on('data', function (chunk) {
            data += chunk;
        });
        res.on("end", function () {
            callback(data);
        });
    }).on("error", function () {
        callback(null);
    });
}

//var cheerio = require("cheerio");

var url = "http://fund.eastmoney.com/data/rankhandler.aspx?op=ph&dt=kf&ft=hh&rs=&gs=0&sc=lnzf&st=desc&sd=2015-08-16&ed=2016-08-16&qdii=&tabSubtype=,,,,,&pi=1&pn=3&dx=1&v=0.37704015895724297";
download(url, function (data) {
    if (data) {
        eval(data);
        let funds = rankData.datas;
        let results = [];
        for (let i = 0, len = funds.length; i < len; i++) {
            let arr = funds[i].split(',');
            results.push(arr);
        }
        saveToExcel(results);
    }
    else {
        console.log("error");
    }
});

//保存到Excel
function saveToExcel(datas) {
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    //一些属性
    workbook.creator = 'cuijie';
    workbook.lastModifiedBy = 'cuijie';
    workbook.created = new Date();
    workbook.modified = new Date();

    var worksheet = workbook.addWorksheet('数据');

    //定义列
    worksheet.columns = [
        { header: '基金代码', key: 'a', width: 10 },
        { header: '成立日', key: 'b', width: 10 },
        { header: '基金简称', key: 'c', width: 10 },
        { header: '单位净值', key: 'd', width: 10 },
        { header: '累计净值', key: 'e', width: 10 },
        { header: '日增长率', key: 'f', width: 10 },
        { header: '近1周', key: 'g', width: 10 },
        { header: '近1月', key: 'h', width: 10 },
        { header: '近3月', key: 'i', width: 10 },
        { header: '近6月', key: 'j', width: 10 },
        { header: '近1年', key: 'k', width: 10 },
        { header: '近2年', key: 'm', width: 10 },
        { header: '近3年', key: 'n', width: 10 },
        { header: '今年来', key: 'o', width: 10 },
        { header: '成立来', key: 'p', width: 10 },
        { header: '基金规模', key: 'q', width: 10 },
        { header: '基金份额', key: 'r', width: 10 },
        { header: '四分位', key: 's', width: 10 },
        { header: '持有人', key: 't', width: 10 },
        { header: '资产配置', key: 'u', width: 10 }
    ];

    let rows = [];
    for (let item of datas) {
        rows.push([item[0], '', item[1], item[4], item[5], item[6], item[7], item[8], item[9],
            item[10], item[11], item[12], item[13], item[14], item[15], '', '', '', '', '']);
    }
    worksheet.addRows(rows);

    workbook.xlsx.writeFile("fund.xlsx")
        .then(function () {
            // done
            console.log("it's done")
        });
}