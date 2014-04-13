var xlsx = require('node-xlsx');
var fs = require('fs');

var config = JSON.parse(fs.readFileSync('./config.json', 'utf-8'));

var excel = xlsx.parse('./test.xlsx');
// console.log("%j", excel);

for (var i_sheet = 0; i_sheet < excel.worksheets.length; i_sheet++) {
    var sheet = excel.worksheets[i_sheet];
    // console.log(sheet.maxCol, sheet.maxRow);
    if (sheet.data && sheet.data.length > 0) {

        var row_head = sheet.data[config.head - 1];

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

            /**
             * 支持的类型
             * 1. basic : number,bool,string
             * 2. object : 自定义数据类型
             * 3. []|[basic] : 普通数据类型数组
             * 4. [object] : 对象数组
             */
            col_type.push(type);
            col_name.push(name);
        };

        //输出json
        var output = [];

        //循环数据行
        for (var i_row = config.head; i_row < sheet.maxRow; i_row++) {
            var row = sheet.data[i_row];

            var row_obj = {}

            //循环一行的每列
            for (var i_col = 0; i_col < sheet.maxCol; i_col++) {
                var cell = row[i_col];

                var type = col_type[i_col].toLowerCase();

                switch (type) {
                    case 'basic':
                        row_obj[col_name[i_col]] = cell.value;
                        break;
                    case 'object':
                        break;
                    case '[]':
                    case '[basic]':
                        if (cell.value) {
                            var basic_array = cell.value.split(',');
                            var temp_array = [];
                            if (isNumberArray(basic_array)) {
                                basic_array.forEach(function(element, index, array) {
                                    temp_array.push(Number(element));
                                });
                            } else if (isBooleanArray(basic_array)) {
                                basic_array.forEach(function(element, index, array) {
                                    temp_array.push(element.toString().toLowerCase() === 'true');
                                });
                            } else { //字符串数组
                                temp_array = basic_array;
                            };
                            // console.log("basic_array", temp_array + "|||" + cell.value);
                            row_obj[col_name[i_col]] = temp_array;
                        };
                        break;
                    case '[bool]':
                        break;
                };
            };

            output.push(row_obj);
        };

        console.log(JSON.stringify(output));
    };
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

    var b = value.toString().toLowerCase();

    return b === 'true' || b === 'false';
};