/**
 * Created by tangyitangyi on 2017/7/31.
 */
const express = require('express');
const XLSX = require('xlsx');
const nodeExcel = require('excel-export');
const path = require('path');
const _ = require('lodash');
const fs = require('fs');
const router = express.Router();

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
/*    console.log(data);*/
    const conf = {};
    const fileName = 'rencaiku';

    conf.name = "aspnetcore";
    conf.cols = [
        {
            caption: '公司(Company)',
            type: 'string',
            width:40
        },
        {
            caption: '行业(Industry)',
            type: 'string',
            width:20
        },
        {
            caption: '姓名(Name)',
            type: 'string',
            width:20
        },
        {
            caption: '手机(Mobile)',
            type: 'string',
            width:20
        },
        {
            caption: '性别(Sex)',
            type: 'string',
            width:20
        },
        {
            caption: '部门(Department)',
            type: 'string',
            width:20
        },
        {
            caption: '职务(JobTitle)',
            type: 'string',
            width:20
        },
        {
            caption: '座机(Telephone)',
            type: 'string',
            width:20
        },
        {
            caption: '电子邮件(Email)',
            type: 'string',
            width:20
        },
        {
            caption: '办公地点(OfficeAddress)',
            type: 'string',
            width:200
        },
        {
            caption: '标签(Tag)',
            type: 'string',
            width:20
        },
        {
            caption: '来源(Source)',
            type: 'string',
            width:20
        }
    ];

    conf.rows = [
/*        ['1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1']*/
    ];
    _.filter(data, function (item, index) {
        var arr01 = [];
        var newObj = {};
        _.map(item, function (obj, clunm) {
            newObj[clunm.replace(/(^\s+)|(\s+$)/g,"").replace(/\s/g,"")] = obj.replace(/(^\s+)|(\s+$)/g,"").replace(/\s/g,"");
        });
        
        if(!_.isEmpty(newObj) && newObj['公司名称'].replace(/(^\s+)|(\s+$)/g,"").replace(/\s/g,"")!=='') {
            arr01[0] = newObj['公司名称'];
            arr01[1] = newObj['经营模式'];
            arr01[2] = newObj['联系人'];
            arr01[3] = newObj['移动电话'];
            arr01[4] = newObj['联系人'].indexOf('先生') > -1 ? '男' : '女';
            arr01[5] = '';
            arr01[6] = '';
            arr01[7] = newObj['电话'];
            arr01[8] = '';
            arr01[9] = newObj['地址'];
            arr01[10] = '';
            arr01[11] = '';
            conf.rows.push(arr01);
            if (index + 1 == 12) {
                arr01 = [];
                newObj = {};
            }
        }
     });

    const result = nodeExcel.execute(conf);

    const random = Math.floor(Math.random() * 10000 + 0);

    const uploadDir = path.join(__dirname, '../writeExcel') + '/';
    const filePath = uploadDir + fileName + '_' + random + ".xlsx";

    fs.writeFile(filePath, result, 'binary', function (err) {
        if (err) {
            console.log(err);
        }

    });
}

const read = (req, res, next) => {

    const workbook = XLSX.readFile(path.join(__dirname, '../readExcel') + '/test.xls');
    console.time("writeExcel");
    writeExcel(to_json(workbook)['1']);
    console.timeEnd("writeExcel");
    res.render('toJson.html', {data: to_json(workbook)});

};

function foo(){
    var x = 4.237;
    var y = 0;
    for (var i=0; i<100000000; i++) {
        y = y + x*x;
    }
    return y;
}

router.get('/', function (req, res, next) {
    read(req, res, next);
});

module.exports = router;