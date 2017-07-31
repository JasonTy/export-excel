/**
 * Created by tangyitangyi on 2017/7/31.
 */
var express = require('express');
var XLSX = require('xlsx');
var nodeExcel = require('excel-export');
var path = require('path');
var _ = require('lodash');
var fs = require('fs');
var router = express.Router();

function to_json(workbook) {
    var result = {};
    // 获取 Excel 中所有表名
    var sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
    workbook.SheetNames.forEach(function (sheetName) {
        var worksheet = workbook.Sheets[sheetName];
        result[sheetName] = XLSX.utils.sheet_to_json(worksheet);
    });
    /*    console.log("打印表信息",JSON.stringify(result, 2, 2));  //显示格式{"表1":[],"表2":[]}*/
    return result;
}

/**
 * 导出excel
 * @param data
 */
const writeExcel = (data) => {
    var conf = {};
    conf.name = "aspnetcore";
    conf.cols = [
        {
            caption: '公司(Company)',
            type: 'string'
        },
        {
            caption: '行业(Industry)',
            type: 'string'
        },
        {
            caption: '姓名(Name)',
            type: 'string'
        },
        {
            caption: '手机(Mobile)',
            type: 'string'
        },
        {
            caption: '性别(Sex)',
            type: 'string'
        },
        {
            caption: '部门(Department)',
            type: 'string'
        },
        {
            caption: '职务(JobTitle)',
            type: 'string'
        },
        {
            caption: '座机(Telephone)',
            type: 'string'
        },
        {
            caption: '电子邮件(Email)',
            type: 'string'
        },
        {
            caption: '办公地点(OfficeAddress)',
            type: 'string'
        },
        {
            caption: '标签(Tag)',
            type: 'string'
        },
        {
            caption: '来源(Source)',
            type: 'string'
        }];

    conf.rows = [
        ['1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1']
    ];

    /*    _.filter(data, function (item) {

     });*/

    var result = nodeExcel.execute(conf);

    var random = Math.floor(Math.random() * 10000 + 0);

    var uploadDir = path.join(__dirname, '../writeExcel') + '/';
    var filePath = uploadDir + random + ".xlsx";

    fs.writeFile(filePath, result, 'binary', function (err) {
        if (err) {
            console.log(err);
        }
    });
}

const read = (req, res, next) => {

    var workbook = XLSX.readFile(path.join(__dirname, '../readExcel') + '/test.xlsx');
    console.log(to_json(workbook));
    writeExcel({});
    res.render('toJson.html', {data: to_json(workbook)});

};

router.get('/', function (req, res, next) {
    read(req, res, next);
});

module.exports = router;