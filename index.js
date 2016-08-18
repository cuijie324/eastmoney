let http = require("http");
let cheerio = require("cheerio");
let fetch = require("node-fetch");
let moment = require("moment");

//开始抓取数据：各种增长率
function start(pagesize) {
    let url = 'http://fund.eastmoney.com/data/rankhandler.aspx?op=ph&dt=kf&ft=hh&rs=&gs=0&sc=lnzf&st=desc&sd=2015-08-16&ed=2016-08-16&qdii=&tabSubtype=,,,,,&pi=1&pn='
        + pagesize + '&dx=1';
    fetch(url)
        .then(function (res) {
            return res.text();
        }).then(function (data) {
            if (data) {
                eval(data);
                let funds = rankData.datas;
                let results = [];
                for (let i = 0, len = funds.length; i < len; i++) {
                    let arr = funds[i].split(',');
                    results.push(arr);
                }

                let rows = [];
                for (let item of results) {
                    rows.push([item[0], '', item[1], item[4], item[5], item[6], item[7], item[8], item[9],
                        item[10], item[11], item[12], item[13], item[14], item[15], '', '', '', '', '']);
                }

                let promises = []
                for (let item of rows) {
                    promises.push(getDetailInfo(item[0]));
                }

                Promise.all(promises).then(datas => {
                    datas.forEach(item => {
                        for (let row of rows) {
                            if (row[0] == item.code) {
                                row[18] = item.holder;
                                row[19] = item.asset;
                            }
                        }
                    })

                    //再从页面里取一些数据，这里有点乱 成立日 基金规模
                    let promise2 = [];
                    for (let item of rows) {
                        promise2.push(getPageInfo(item[0]));
                    }
                    Promise.all(promise2).then(datas => {
                        datas.forEach(item => {
                            for (let row of rows) {
                                if (row[0] == item.code) {
                                    row[1] = item.result[1];
                                    row[15] = item.result[0];
                                }
                            }
                        });

                        //再从页面里取一些数据，这里有点乱
                        let promise3 = [];
                        for (let item of rows) {
                            promise3.push(getPageInfo2(item[0]));
                        }
                        Promise.all(promise3).then(datas => {
                            datas.forEach(item => {
                                for (let row of rows) {
                                    if (row[0] == item.code) {
                                        row[16] = item.ab;
                                    }
                                }
                            });

                            //再从页面里取一些数据，这里有点乱
                            let promise4 = [];
                            for (let item of rows) {
                                promise4.push(getPageInfo3(item[0]));
                            }
                            Promise.all(promise4).then(datas => {
                                datas.forEach(item => {
                                    for (let row of rows) {
                                        if (row[0] == item.code) {
                                            row[17] = item.num;
                                        }
                                    }
                                });

                                saveToExcel(rows);
                            });
                        });
                    });
                }).catch(err => console.error(err));
            }
            else {
                console.log("error");
            }
        }).catch(err => console.error(err));
}

start(3);

//获取每只基金的数据 http://fund.eastmoney.com/pingzhongdata/000011.js
function getDetailInfo(code) {
    return new Promise(function (resolve, reject) {
        let url = 'http://fund.eastmoney.com/pingzhongdata/' + code + '.js';
        console.log(url);
        fetch(url)
            .then(function (res) {
                return res.text();
            }).then(function (data) {
                eval(data);
                //资产配置 Data_assetAllocation            
                let asset = '';
                for (let item of Data_assetAllocation.series) {
                    asset += item.name.slice(0, 2) + '=' + item.data.pop().toFixed(2) + '% ';
                }
                asset = asset.trim().slice(0, -1) + '亿元';

                //持有人结构 Data_holderStructure
                let holder = '';
                for (let item of Data_holderStructure.series) {
                    holder += item.name.slice(0, 2) + '=' + item.data.pop().toFixed(2) + '% ';
                }
                resolve({ code, asset, holder });

            }).catch(err => reject(err));
    });
}

//抓取页面数据：成立日、基金规模 http://fund.eastmoney.com/000011.html
function getPageInfo(code) {
    return new Promise(function (resolve, reject) {
        let url = 'http://fund.eastmoney.com/' + code + '.html';
        console.log(url);
        fetch(url)
            .then(function (res) {
                return res.text();
            }).then(function (body) {
                let $ = cheerio.load(body);
                let result = [];
                $('.merchandiseDetail .infoOfFund table tr td').each(function (i, e) {
                    var err = $(e);
                    if (i == 1) {
                        result.push(err.text().trim().slice(5));
                    }

                    if (i == 3) {
                        result.push(err.text().trim().slice(-10));
                    }
                });
                console.log(result);
                resolve({ code, result });
            }).catch(err => reject(err));
    });
}

//getPageInfo2('000011');

//抓取页面数据：份额规模 http://fund.eastmoney.com/f10/jbgk_000011.html
function getPageInfo2(code) {
    return new Promise(function (resolve, reject) {
        let url = 'http://fund.eastmoney.com/f10/jbgk_' + code + '.html';
        fetch(url)
            .then(function (res) {
                return res.text();
            }).then(function (body) {
                let $ = cheerio.load(body);
                $('.txt_cont table tr').each(function (i, e) {
                    if (i == 3) {
                        var ab = $(e).find('a').text();
                        resolve({ code, ab });
                    }
                });
            }).catch(err => reject(err));
    });
}

//getPageInfo3('000011');

//抓取页面数据：四分位 http://fund.eastmoney.com/f10/FundArchivesDatas.aspx?type=jdzf&code=000011
function getPageInfo3(code) {
    return new Promise(function (resolve, reject) {
        let url = 'http://fund.eastmoney.com/f10/FundArchivesDatas.aspx?type=jdzf&code=' + code;
        fetch(url)
            .then(function (res) {
                return res.text();
            }).then(function (body) {
                eval(body);
                let num = 0;
                let $ = cheerio.load(apidata.content);
                $('li.sf').each(function (i, e) {
                    if ($(e).text().trim() == '优秀') {
                        num++;
                    }
                });
                resolve({ code, num });
            }).catch(e => resolve(err));
    });
}

//保存到Excel
function saveToExcel(rows) {
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
        { header: '持有人', key: 't', width: 30 },
        { header: '资产配置', key: 'u', width: 30 }
    ];

    worksheet.addRows(rows);

    let filename = 'fund_' + moment().format('YYYY-MM-DD') + '.xlsx';
    workbook.xlsx.writeFile(filename)
        .then(function () {
            // done
            console.log("it's done")
        });
}