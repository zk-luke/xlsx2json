const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');

let _config;

function parseJsonObject(data) {
  var evil = eval;
  return evil("(" + data + ")");
}


// class StringBuffer {
//   constructor(str) {
//     this._str_ = [];
//     if (str) {
//       this.append(str);
//     }
//   }

//   toString() {
//     return this._str_.join("");
//   }

//   append(str) {
//     this._str_.push(str);
//   }
// }

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
  UNKOWN: 'unkown'
};

/**
 * convert worksheet in workbook and save to file for each.
 * @param       {[Object]} workbook json object of excel's workbook.
 * @param       {[String]} dest     directory for exported json files.
 * @param       {[Number]} headIndex     index of head line.
 */
function parseWorkbook(workbook, dest, headIndex, excelName) {

  let dtsstring = "";

  workbook.forEach(sheet => {

    // ignore sheet with external keys, or start with a '!'
    if (sheet.name.indexOf('@') === -1 && !sheet.name.startsWith('!')) {
      let parsedSheet = parseSheet(sheet, headIndex);

      let dest_file = path.resolve(dest, sheet.name + ".json");
      let resultJson = JSON.stringify(parsedSheet.result, null, _config.json.uglify ? 0 : 2); //, null, 2
      fs.writeFile(dest_file, resultJson, err => {
        if (err) {
          console.error("error：", err);
          throw err;
        }
        console.log('exported json successfully  -->  ', path.basename(dest_file));
      });

      if (_config.ts) {
        dtsstring += formatDTS(sheet.name, parsedSheet.head);
      }
    }
  });

  if (_config.ts) {
    let dest_file_dts = path.resolve(dest, excelName + ".d.ts");
    fs.writeFile(dest_file_dts, dtsstring, err => {
      if (err) {
        console.error("error：", err);
        throw err;
      }
      console.log('exported t.ds successfully  -->  ', path.basename(dest_file_dts));
    });
  }
}

/**
 * 
 * @param {String} name the excel file name will be use on create d.ts
 * @param {Object} head the excel head will be the javescript field
 */
function formatDTS(name, head) {
  let className = name.substring(0, 1).toUpperCase() + name.substring(1);
  let strHead = "interface " + className + "Template {\r\n";
  for (let i = 0; i < head.names.length; ++i) {
    let typesDes = "any";
    switch (head.types[i]) {
      case DataType.NUMBER:
        {
          typesDes = "number";
          break;
        }
      case DataType.STRING:
        {
          typesDes = "string";
          break;
        }
      case DataType.BOOL:
        {
          typesDes = "boolean";
          break;
        }
      case DataType.ID:
        {
          typesDes = "string";
          break;
        }
      case DataType.ARRAY:
        {
          typesDes = "any[]";
          break;
        }
      case DataType.OBJECT:
        {
          typesDes = "any";
          break;
        }
      case DataType.UNKOWN:
        {
          typesDes = "any";
          break;
        }
      default:
        {
          typesDes = "any";
        }
    }
    strHead += "\t" + head.names[i] + ": " + typesDes + "\r\n";
  }
  strHead += "}\r\n";
  return strHead;
}

/**
 * parse one sheet and return the result as a json object or array
 *
 * @param sheet
 * @param headIndex
 * @private
 */
function parseSheet(sheet, headIndex) {

  console.log('  * sheet:', sheet.name);

  if (sheet && sheet.data) {

    let head = parseHead(sheet, headIndex);

    // console.log('\t parsing head', JSON.stringify(head));

    let result;

    if (head.sheetType === SheetType.NORMAL) {
      result = [];
    } else if (head.sheetType === SheetType.PRIMARY) {
      result = {};
    }

    for (let i_row = headIndex + 1; i_row < sheet.data.length; i_row++) {

      let row = sheet.data[i_row];

      let parsedRow = parseRow(row, i_row, head);
      if (head.sheetType === SheetType.NORMAL) { // json以数组的格式输出
        result.push(parsedRow);
      } else if (head.sheetType === SheetType.PRIMARY) { //json以map的格式输出
        let id = parsedRow[head.getIdKey()];
        result[id] = parsedRow;
      } else {
        throw '无法识别表格类型!';
      }
    }
    return {
      result,
      head
    };
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

    getIdKey: function () {
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

  // console.log('parsing row', row);

  for (let index = 0; index < head.names.length; index++) {
    let cell = row[index];

    let name = head.names[index];
    let type = head.types[index];

    if (name.startsWith('!')) {
      continue;
    }

    if (!cell) {
      result[name] = null;
      continue;
    }

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
          result[name] = cell;
        }
        break;
      case DataType.DATE:
        if (isNumber(cell)) {
          //xlsx's bug!!!
          result[name] = numdate(cell);
        } else {
          result[name] = cell.toString();
        }
        break;
      case DataType.STRING:
        if (cell.startsWith('"')) {
          result[name] = parseJsonObject(cell);
        } else {
          result[name] = cell.toString();
        }
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
      case DataType.OBJECT:
        result[name] = parseJsonObject(cell);
        break;
      case DataType.ARRAY:
        result[name] = parseJsonObject(`[${cell}]`);
        break;
      default:
        console.log('无法识别的类型:', '[' + rowIndex + ',' + index + ']', cell, typeof (cell));
        break;
    }
  }

  return result;
}

/**
 * convert value to boolean.
 */
function toBoolean(value) {
  return value.toString().toLowerCase() === 'true';
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

  if (typeof (value) === "undefined") {
    return false;
  }

  if (typeof value === 'boolean') {
    return true;
  }

  let b = value.toString().trim().toLowerCase();

  return b === 'true' || b === 'false';
}

//fuck node-xlsx's bug
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
  toJson: function (src, dest, config) {

    _config = config;

    let headIndex = config.xlsx.head - 1;

    if (!fs.existsSync(dest)) {
      fs.mkdirSync(dest);
    }

    let parsed_src = path.parse(src);

    let workbook = xlsx.parse(src);

    console.log("parsing excel:", parsed_src.base);

    parseWorkbook(workbook, dest, headIndex, path.join(dest, parsed_src.name));
  }
};