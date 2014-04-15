'use strict'

var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');

// console.log("%j", excel);

module.exports = {
    parse: function(config) {

        fs.readdir(config.src, function(err, files) {
            if (err) {
                console.error("error", err);
                throw err;
            };

            var filtered = files.filter(function(element) {
                return path.extname(element) === '.xlsx';
            });

            //todo:性能瓶颈
            filtered.forEach(function(element, index, array) {
                var excel = xlsx.parse(path.resolve(config.src, element));
                toJson(excel, config.head, config.dest);
            });
        });
    }
};


/**
 * 将原始excel json数据转换为我们需要的json格式。
 * excel:node-xlsx 导出的json格式。
 * head : 表头所在的行号
 * dest : 目标文件夹
 */
function toJson(excel, head, dest) {
    for (var i_sheet = 0; i_sheet < excel.worksheets.length; i_sheet++) {
        var sheet = excel.worksheets[i_sheet];
        // console.log(sheet.maxCol, sheet.maxRow);
        if (sheet.data && sheet.data.length > 0) {

            var row_head = sheet.data[head - 1];

            var col_type = []; //列 类型
            var col_name = []; //列 名字

            //读取表头 解读列名字和列数据类型
            for (var i_cell = 0; i_cell < row_head.length; i_cell++) {
                var name = row_head[i_cell].value;
                var type = 'basic';

                if (name.indexOf('#') != -1) {
                    var temp = name.split('#');
                    name = temp[0];
                    type = temp[1];
                };

                col_type.push(type);
                col_name.push(name);
            };

            //输出json
            var output = [];

            //循环数据行
            for (var i_row = head; i_row < sheet.maxRow; i_row++) {
                var row = sheet.data[i_row];

                var row_obj = {}

                //循环一行的每列
                for (var i_col = 0; i_col < sheet.maxCol; i_col++) {
                    var cell = row[i_col];

                    var type = col_type[i_col].toLowerCase().trim();
                    // console.log("type:", type);
                    switch (type) {
                        case 'basic': // number string boolean
                            row_obj[col_name[i_col]] = cell.value;
                            break;
                        case '{}':
                            if (cell) {
                                addObjectField(row_obj, col_name[i_col], cell.value);
                            };
                            break;
                        case '[]': //[number] [boolean] [string]
                            if (cell) {
                                addBasicArrayField(row_obj, col_name[i_col], cell.value);
                            };
                            break;
                        case '[{}]': //[object]
                            if (cell) {
                                addObjectArrayField(row_obj, col_name[i_col], cell.value);
                            };
                            break;
                    };
                };

                output.push(row_obj);
            };

            output = JSON.stringify(output);
            if (!fs.existsSync(dest)) {
                fs.mkdirSync(dest);
            };
            var dest_file = path.resolve(dest, sheet.name + ".json");
            fs.writeFile(dest_file, output, function(err) {
                if (err) {
                    console.log("error", err);
                    throw err;
                }
                console.log('已导出  -->  ', path.basename(dest_file));
            });
        };
    };
};


/**
 * 拼装对象数组
 */
function addObjectArrayField(field, key, array) {
    var obj_array = array.split(',');

    var result = [];

    obj_array.forEach(function(element, index, array) {
        result.push(array2object(element.split(';')));
    });

    field[key] = result;
};

/**
 * 从数组创建对象
 * 比如 [a:123,b:45] => {'a':123,'b':45}
 */
function array2object(array) {
    var obj_field = array;
    var result = {};
    obj_field.forEach(function(element, index, array) {
        var kv = element.split(':');
        if (isNumber(kv[1])) {
            kv[1] = Number(kv[1]);
        } else if (isBoolean(kv[1])) {
            kv[1] = toBoolean(kv[1]);
        };
        result[kv[0]] = kv[1];
    });
    return result;
}

/**
 * 拼装对象
 */
function addObjectField(field, key, data) {
    field[key] = array2object(data.split(';'));
};


/**
 * 拼装普通数组
 */
function addBasicArrayField(field, key, array) {
    var basic_array = array.split(',');
    var result = [];
    if (isNumberArray(basic_array)) {
        basic_array.forEach(function(element, index, array) {
            result.push(Number(element));
        });
    } else if (isBooleanArray(basic_array)) {
        basic_array.forEach(function(element, index, array) {
            result.push(toBoolean(element));
        });
    } else { //字符串数组
        result = basic_array;
    };
    // console.log("basic_array", result + "|||" + cell.value);
    field[key] = result;
};

/**
 * 将一个true false 形式的字符串或者其他什么的变量转成boolean
 *
 * 不能直接用Boolean构造，任何非空的字符串都会被当作 true，具体见如下几种情况。
 * var myBool = Boolean("false");  // == true
 * var myBool = !!"false";  // == true
 * myFalse = new Boolean(false);   // initial value of false
 * g = Boolean(myFalse);       // initial value of true
 * 当比较的变量是不同的数据类型时候 使用 === 来代替 ==。
 */
function toBoolean(value) {
    return value.toString().toLowerCase() === 'true'
};

/**
 * 判断一个数组是不是bool数组
 */
function isBooleanArray(arr) {
    return arr.every(function(element, index, array) {
        return isBoolean(element);
    });
};

/**
 * 判断一个数组是不是数字数组
 */
function isNumberArray(arr) {
    return arr.every(function(element, index, array) {
        return isNumber(element);
    });
};

/**
 * 判断一个变量是不是数字类型
 */
function isNumber(value) {
    if (typeof value === 'number') {
        return true;
    };
    return !isNaN(+value.toString());
};

/**
 * 判断一个变量是不是bool类型
 */
function isBoolean(value) {
    if (typeof value === 'boolean') {
        return true;
    };

    var b = value.toString().trim().toLowerCase();

    return b === 'true' || b === 'false';
};