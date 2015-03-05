'use strict'

var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');


var arraySeparator;

module.exports = {

    /**
     * export .xlsx file to json
     * src_excel_file: path of .xlsx files.
     * dest_dir: directory for exported json files.
     * head: line number of excell headline.
     * separator : array separator.
     */
    toJson: function(src_excel_file, dest_dir, head, separator) {

        arraySeparator = separator;

        if (!fs.existsSync(dest_dir)) {
            fs.mkdirSync(dest_dir);
        };
        console.log("parsing excel:", src_excel_file);
        var excel = xlsx.parse(src_excel_file);
        _toJson(excel, dest_dir, head);
    }
};


/**
 * export .xlsx file to json formate.
 * excel: json string converted by 'node-xlsx'。
 * head : line number of excell headline.
 * dest : directory for exported json files.
 */
function _toJson(excel, dest, head) {
    for (var i_sheet = 0; i_sheet < excel.worksheets.length; i_sheet++) {
        var sheet = excel.worksheets[i_sheet];
        console.log("cols:" + sheet.maxCol, "rows:" + sheet.maxRow);
        if (sheet.data && sheet.data.length > 0) {

            var row_head = sheet.data[head - 1];

            var col_type = []; //column data type
            var col_name = []; //column name

            //读取表头 解读列名字和列数据类型
            //parse headline to get column name & column data type
            for (var i_cell = 0; i_cell < row_head.length; i_cell++) {
                var name = row_head[i_cell].value;
                if(typeof name == 'undefined' || !name){
                    break;
                }

                var type = 'basic';

                if (name.indexOf('#') != -1) {
                    var temp = name.split('#');
                    name = temp[0];
                    type = temp[1];
                };

                col_type.push(type);
                col_name.push(name);
            };

            var output = [];

            for (var i_row = head; i_row < sheet.maxRow; i_row++) {
                var row = sheet.data[i_row];
                if(typeof row == 'undefined' || !row[0]){
                    break;
                }
                var json_obj = {}

                for (var i_col = 0; i_col < col_type.length; i_col++) {
                    var cell = row[i_col];
                    var type = col_type[i_col].toLowerCase().trim();

                    switch (type) {
                        case 'basic': // number string boolean
                            if (cell) {
                                json_obj[col_name[i_col]] = cell.value;
                            };
                            break;
                        case 'string':
                            if (cell && !isNaN(+cell.value.toString())) {
                                json_obj[col_name[i_col]] = cell.value.toString();
                            }else{
                                json_obj[col_name[i_col]]="";
                            };
                            break;
                        case 'number':
                            if (cell) {
                                json_obj[col_name[i_col]] = Number(cell.value);
                            }else{
                                json_obj[col_name[i_col]]=0;
                            };
                            break;
                        case 'bool':
                            if (cell) {
                                json_obj[col_name[i_col]] = toBoolean(cell.value.toString());
                            }else{
                                json_obj[col_name[i_col]]=false;
                            };
                            break;
                        case '{}':
                            if (cell) {
                                parseObject(json_obj, col_name[i_col], cell.value);
                            };
                            break;
                        case '[]': //[number] [boolean] [string]
                            if (cell) {
                                parseBasicArrayField(json_obj, col_name[i_col], cell.value);
                            };
                            break;
                        case '[{}]': //[object]
                            if (cell) {
                                parseObjectArrayField(json_obj, col_name[i_col], cell.value);
                            }else{
                                json_obj[col_name[i_col]]=[];
                            };
                            break;
                    };
                };

                output.push(json_obj);
            };

            output = JSON.stringify(output, null, 2);

            var dest_file = path.resolve(dest, sheet.name + ".json");
            fs.writeFile(dest_file, output, function(err) {
                if (err) {
                    console.log("error：", err);
                    throw err;
                }
                console.log('exported successfully  -->  ', path.basename(dest_file));
            });
        };
    };
};


/**
 * parse object array.
 */
function parseObjectArrayField(field, key, array) {

    var obj_array;

    if (typeof(array) === 'string' && array.indexOf(arraySeparator) !== -1) {
        obj_array = array.split(arraySeparator);
    } else {
        obj_array = [];
        obj_array.push(array);
    };

    var result = [];

    obj_array.forEach(function(element, index, array) {
        if (element) {
            result.push(array2object(element.split(';')));
        }
    });

    field[key] = result;
};

/**
 * parse object from array.
 *  for example : [a:123,b:45] => {'a':123,'b':45}
 */
function array2object(array) {
    var obj_field = array;
    var result = {};
    obj_field.forEach(function(element, index, array) {
        if (element) {
            var kv = element.trim().split(':');
            if (isNumber(kv[1])) {
                kv[1] = Number(kv[1]);
            } else if (isBoolean(kv[1])) {
                kv[1] = toBoolean(kv[1]);
            };
            result[kv[0]] = kv[1];
        };
    });
    return result;
}

/**
 * parse object
 */
function parseObject(field, key, data) {
    field[key] = array2object(data.split(';'));
};


/**
 * parse simple array.
 */
function parseBasicArrayField(field, key, array) {
    var basic_array;

    if (typeof array === "string") {
        basic_array = array.split(arraySeparator);
    } else {
        basic_array = [];
        basic_array.push(array);
    };

    var result = [];
    if (isNumberArray(basic_array)) {
        basic_array.forEach(function(element, index, array) {
            result.push(Number(element));
        });
    } else if (isBooleanArray(basic_array)) {
        basic_array.forEach(function(element, index, array) {
            result.push(toBoolean(element));
        });
    } else { //string array
        result = basic_array;
    };
    // console.log("basic_array", result + "|||" + cell.value);
    field[key] = result;
};

/**
 * convert value to boolean.
 */
function toBoolean(value) {
    return value.toString().toLowerCase() === 'true'
};

/**
 * is a boolean array.
 */
function isBooleanArray(arr) {
    return arr.every(function(element, index, array) {
        return isBoolean(element);
    });
};

/**
 * is a number array.
 */
function isNumberArray(arr) {
    return arr.every(function(element, index, array) {
        return isNumber(element);
    });
};

/**
 * is a number.
 */
function isNumber(value) {

    if (typeof(value) == "undefined") {
        return false;
    }

    if (typeof value === 'number') {
        return true;
    };
    return !isNaN(+value.toString());
};

/**
 * is boolean type.
 */
function isBoolean(value) {

    if (typeof(value) == "undefined") {
        return false;
    }

    if (typeof value === 'boolean') {
        return true;
    };

    var b = value.toString().trim().toLowerCase();

    return b === 'true' || b === 'false';
};