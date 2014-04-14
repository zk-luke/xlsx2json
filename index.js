var xlsx = require('./lib/xlsx-to-json.js');
var path = require('path');

var config = {
    "head": 2, //表头所在的行
    "src": path.join(__dirname, "excel"), //excel目录。
    "dest": path.join(__dirname, "json") //目标目录
};

xlsx.parse(config);