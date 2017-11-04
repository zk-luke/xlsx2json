const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');
// const moment = require('moment');

let arraySeparator;

/**
 * sheet(table) 的类型
 * 影响输出json的类型
 * 当有#id类型的时候  表输出json的是map形式(id:{xx:1})
 * 当没有#id类型的时候  表输出json的是数组类型 没有id索引
 */
const SheetType = {
  /**
   * 普通表
   */
  NORMAL: 0,

  /**
   * 有主外键关系的主表
   * primary key
   */
  PRIMARY: 1,

  /**
   * 有主外键关系的附表
   * foreign key
   */
  FOREIGN: 2
};

/**
 * 支持的数据类型
 */
const DataType = {
  NUMBER: 'number',
  STRING: 'string',
  BOOL: 'bool',
  DATE: 'date',
  ID: 'id',
  ARRAY: '[]',
  OBJECT: '{}',
  OBJECT_ARRAY: '[{}]',
  UNKOWN: 'unkown'
};

module.exports = {

  /**
   * convert xlsx file to json and save it to file system.
   * @param  {String} src path of .xlsx files.
   * @param  {String} dest       directory for exported json files.
   * @param  {Number} headIndex      index of head line.
   * @param  {String} separator      array separator.
   *
   * excel structure
   * workbook > worksheet > table(row column)
   */
  toJson: function(src, dest, config) {

    let headIndex = config.xlsx.head - 1;
    arraySeparator = config.xlsx.arraySeparator;
    let uglifyJson = config.json.uglify;

    if (!fs.existsSync(dest)) {
      fs.mkdirSync(dest);
    }

    console.log("parsing excel:", src);

    let workbook = xlsx.parse(src);
    parseWorkbook(workbook, dest, headIndex, uglifyJson);
  }
};

/**
 * convert worksheet in workbook and save to file for each.
 * @param       {[Object]} workbook json object of excel's workbook.
 * @param       {[String]} dest     directory for exported json files.
 * @param       {[Number]} headIndex     index of head line.
 */
function parseWorkbook(workbook, dest, headIndex, uglifyJson) {

  workbook.forEach(sheet => {

    // ignore sheet with external keys only, or start with a '#'
    if (sheet.name.indexOf('@') === -1 && sheet.name[0] !== "#") {
      let parsedSheet = parseSheet(sheet, headIndex);

      let dest_file = path.resolve(dest, sheet.name + ".json");
      let formatedJson = JSON.stringify(parsedSheet, null, uglifyJson ? 0 : 2); //, null, 2
      fs.writeFile(dest_file, formatedJson, err => {
        if (err) {
          console.error("error：", err);
          throw err;
        }
        console.log('exported successfully  -->  ', path.basename(dest_file));
      });
    }
  });
}

/**
 * parse one sheet and return the result as a json object or array
 *
 * @param sheet
 * @param headIndex
 * @private
 */
function parseSheet(sheet, headIndex) {

  console.log('\t parsing sheet', sheet.name);

  if (sheet && sheet.data) {

    let head = parseHead(sheet, headIndex);

    let result;

    if (head.sheetType === SheetType.NORMAL) {
      result = [];
    } else if (head.sheetType === SheetType.PRIMARY) {
      result = {};
    }

    for (let i_row = headIndex + 1; i_row < sheet.data.length; i_row++) {

      let row = sheet.data[i_row];

      let parsedRow = parseRow(row, i_row, head);
      if (head.sheetType === SheetType.NORMAL) { ////json以数组的格式输出
        result.push(parsedRow);
      } else if (head.sheetType === SheetType.PRIMARY) { //json以map的格式输出
        let id = parsedRow[head.getIdKey()];
        result[id] = parsedRow;
      } else {
        throw '无法识别表格类型!';
      }
    }
    return result;
  }
}

function parseHead(sheet, headIndex) {

  let headRow = sheet.data[headIndex];

  // console.log("\t\t parsing head", headIndex, headRow);

  let head = {
    //所有名字 json的key
    names: [],

    //所有列的数据类型
    types: [],

    //表头所在的行索引
    index: headIndex,

    //表类型 普通表  主表 引用表
    sheetType: SheetType.NORMAL,

    getIdKey: function() {
      let id_col_index = this.types.indexOf(DataType.ID);
      if (id_col_index < 0) {
        throw '获取不到id列的名字';
      }
      return this.names[id_col_index];
    }
  };

  headRow.forEach(cell => {
    let type = DataType.UNKOWN;
    let name = cell;

    if ((cell + '').indexOf('#') !== -1) {
      let pair = cell.split('#');
      name = pair[0].trim();
      type = pair[1].toLowerCase().trim();
      if (type === DataType.ID) {
        head.sheetType = SheetType.PRIMARY;
      }
    }

    head.types.push(type);
    head.names.push(name);
  });

  return head;
}

function parseRow(row, rowIndex, head) {

  let result = {};
  let id;

  console.log('parsing row', row);

  row.forEach((cell, index) => {

    // if (cell) {

    let name = head.names[index];
    let type = head.types[index];

    switch (type) {
      case DataType.ID: // number string boolean
        if (isNumber(cell)) {
          id = Number(cell);
        } else {
          id = cell;
        }
        result[name] = id;
        break;
      case DataType.UNKOWN: // number string boolean
        if (isNumber(cell)) {
          result[name] = Number(cell);
        } else if (isBoolean(cell)) {
          result[name] = toBoolean(cell);
        } else {
          if (cell) {
            result[name] = cell;
          }
        }
        break;
      case DataType.DATE:
        if (isNumber(cell)) {
          //xlsx's bug!!!
          result[name] = numdate(cell);
        } else {
          if (cell) {
            result[name] = cell.toString();
          }
        }
        break;
      case DataType.STRING:
        result[name] = cell.toString();
        break;
      case DataType.NUMBER:
        //+xxx.toString() '+' means convert it to number
        if (isNumber(cell)) {
          result[name] = Number(cell);
        } else {
          console.warn("type error at [" + rowIndex + "," + index + "]," + cell + " is not a number");
        }
        break;
      case DataType.BOOL:
        result[name] = toBoolean(cell);
        break;
      case DataType.OBJECT: //support {number boolean string date} property type
        if (cell) {
          result[name] = array2object(cell.split(';'));
        }
        break;
      case DataType.ARRAY: //[number] [boolean] [string]  todo:support [date] type
        result[name] = parseBasicArrayField(cell, arraySeparator);
        break;
      case DataType.OBJECT_ARRAY:
        result[name] = parseObjectArrayField(cell);
        break;
      default:
        // foo#[]| 处理自定义数组分隔符
        if (type.indexOf(DataType.ARRAY) !== -1) {
          // if (!type.endsWith(DataType.ARRAY)) {
          let separator = type.substr(-1, 1); //get the last character
          result[name] = parseBasicArrayField(cell, separator);
          // }
        } else {
          console.log('unrecognized type', '[' + rowIndex + ',' + index + ']', cell, typeof(cell));
        }
        break;
    }
    // }
  });

  return result;

  // switch (head.sheetType) {
  //   case SheetType.NORMAL: //json以数组的格式输出
  //     return result;
  //   case SheetType.PRIMARY: //json以map的格式输出
  //     let map = {};
  //     map[id] = result;
  //     return map;
  //   default:
  //     throw '无法识别表格类型!';
  // }

  // return result;
}

/**
 * parse object array.
 */
function parseObjectArrayField(value) {

  let obj_array = [];

  if (value) {
    if (value.indexOf(',') !== -1) {
      obj_array = value.split(',');
    } else {
      obj_array.push(value.toString());
    }
  }

  // if (typeof(value) === 'string' && value.indexOf(',') !== -1) {
  //     obj_array = value.split(',');
  // } else {
  //     obj_array.push(value.toString());
  // };

  let result = [];

  obj_array.forEach(function(e) {
    if (e) {
      result.push(array2object(e.split(';')));
    }
  });

  // row[key] = result;
  return result;
}

/**
 * parse object from array.
 *  for example : [a:123,b:45] => {'a':123,'b':45}
 */
function array2object(array) {
  let result = {};
  array.forEach(function(e) {
    if (e) {
      let colonIndex = e.indexOf(':');
      let key = e.substring(0, colonIndex);
      let value = e.substring(colonIndex + 1);
      if (isNumber(value)) {
        value = Number(value);
      } else if (isBoolean(value)) {
        value = toBoolean(value);
      }
      result[key] = value;
    }
  });
  return result;
}

/**
 * parse simple array.
 */
function parseBasicArrayField(array, arraySeparator) {
  let basic_array;

  if (typeof array === "string") {
    basic_array = array.split(arraySeparator);
  } else {
    basic_array = [];
    basic_array.push(array);
  }

  let result = [];
  if (isNumberArray(basic_array)) {
    basic_array.forEach(function(element) {
      result.push(Number(element));
    });
  } else if (isBooleanArray(basic_array)) {
    basic_array.forEach(function(element) {
      result.push(toBoolean(element));
    });
  } else { //string array
    result = basic_array;
  }
  // console.log("basic_array", result + "|||" + cell.value);
  // field[key] = result;
  return result;
}

/**
 * convert value to boolean.
 */
function toBoolean(value) {
  return value.toString().toLowerCase() === 'true';
}

/**
 * is a boolean array.
 */
function isBooleanArray(arr) {
  return arr.every(function(element, index, array) {
    return isBoolean(element);
  });
}

/**
 * is a number array.
 */
function isNumberArray(arr) {
  return arr.every(function(element, index, array) {
    return isNumber(element);
  });
}

/**
 * is a number.
 */
function isNumber(value) {

  if (typeof value === 'number') {
    return true;
  }

  if (value) {
    return !isNaN(+value.toString());
  }

  return false;
}

/**
 * boolean type check.
 */
function isBoolean(value) {

  if (typeof(value) === "undefined") {
    return false;
  }

  if (typeof value === 'boolean') {
    return true;
  }

  let b = value.toString().trim().toLowerCase();

  return b === 'true' || b === 'false';
}

/**
 * date type check.
 */
// function isDateType(value) {
//   if (isNumber(value)) {
//     return false;
//   }
//   if (value) {
//     return moment(new Date(value), "YYYY-M-D", true).isValid() || moment(value, "YYYY-M-D H:m:s", true).isValid() || moment(value, "YYYY/M/D H:m:s", true).isValid() || moment(value, "YYYY/M/D", true).isValid();
//   }
//   return false;
// }

//fuck xlsx's bug
var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
// var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
// function datenum(v, date1904) {
// 	var epoch = v.getTime();
// 	if(date1904) epoch -= 1462*24*60*60*1000;
// 	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
// }

function numdate(v) {
  var out = new Date();
  out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
  return out;
}
//fuck over
