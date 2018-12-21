const config = require('../config.json');
const types = require('./types');
var _ = require('lodash');

const DataType = types.DataType;
const SheetType = types.SheetType;

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
 * 解析workbook中所有sheet的设置
 * @param {*} workbook 
 */
function parseSettings(workbook) {

  /**
   * settings's schema
   * {
      type: SheetType.NORMAL,
      head: [{
          name: "json-key(column name)",
          type: "number/string/bool/data/id/[]/{}",
        }]
      ,
      slaves: [],
      master: {
        name: '',
        type: DataType.OBJECT
      },
    } 
   */

  let settings = {};

  workbook.forEach(sheet => {

    //叹号开头的sheet不输出
    if (sheet.name.startsWith('!')) {
      return;
    }

    let slave = sheet.name.indexOf('@') >= 0;
    let sheet_name = sheet.name;

    let sheet_setting = {
      type: SheetType.NORMAL,
      master: null,
      slaves: [],
      head: []
    };

    if (slave) {
      sheet_setting.type = SheetType.SLAVE;
      let pair = sheet_name.split('@');
      sheet_name = pair[0].trim();
      sheet_setting.master = pair[1].trim();
      settings[sheet_setting.master].slaves.push(sheet_name);
    }

    let head_row = sheet.data[config.xlsx.head - 1];

    //parsing head setting
    head_row.forEach(cell => {

      cell = cell.toString();

      let head_setting = {
        name: cell,
        type: DataType.UNKNOWN,
      };

      if (cell.indexOf('#') !== -1) {
        let pair = cell.split('#');
        let name = pair[0].trim();
        let type = pair[1].trim();

        head_setting.name = name;
        head_setting.type = type;

        if (!slave && type === DataType.ID) {
          sheet_setting.type = SheetType.MASTER;
        }
      }
      sheet_setting.head.push(head_setting);
    });

    settings[sheet_name] = sheet_setting;
  });

  return settings;
}


/**
 * 解析一个表(sheet)
 *
 * @param sheet 表的原始数据
 * @param setting 表的设置
 * @return Array or Object
 */
function parseSheet(sheet, setting) {

  let headIndex = config.xlsx.head;
  let result = [];

  console.log('  * sheet:', sheet.name, 'rows:', sheet.data.length);

  if (setting.type === SheetType.MASTER) {
    result = {};
  }

  for (let i_row = headIndex; i_row < sheet.data.length; i_row++) {

    let row = sheet.data[i_row];

    let parsed_row = parseRow(row, i_row, setting.head);

    if (setting.type === SheetType.MASTER) {

      let id_cell = _.find(setting.head, item => {
        return item.type === DataType.ID || item.type === DataType.IDS;
      });

      if (!id_cell) {
        throw `在表${sheet.name}中获取不到id列`;
      }

      result[parsed_row[id_cell.name]] = parsed_row;

    } else {
      result.push(parsed_row);
    }
  }

  return result;
}

/**
 * 解析一行
 * @param {*} row 
 * @param {*} rowIndex 
 * @param {*} head 
 */
function parseRow(row, rowIndex, head) {

  let result = {};
  let id;

  for (let index = 0; index < head.length; index++) {
    let cell = row[index];

    let name = head[index].name;
    let type = head[index].type;

    if (name.startsWith('!')) {
      continue;
    }

    if (cell === null || cell === undefined) {
      result[name] = null;
      continue;
    }

    switch (type) {
      case DataType.ID:
        id = cell + '';
        result[name] = id;
        break;
      case DataType.IDS:
        id = cell + '';
        result[name] = id;
        break;
      case DataType.UNKNOWN: // number string boolean
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
        if (cell.toString().startsWith('"')) {
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
        if (!cell.toString().startsWith('[')) {
          cell = `[${cell}]`;
        }
        result[name] = parseJsonObject(cell);
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

  parseSettings: parseSettings,

  parseWorkbook: function (workbook, settings) {

    // console.log('settings >>>>>', JSON.stringify(settings, null, 2));

    let parsed_workbook = {};

    workbook.forEach(sheet => {

      if (sheet.name.startsWith('!')) {
        return;
      }

      let sheet_name = sheet.name;

      let isSlave = sheet_name.indexOf('@') >= 0;

      if (isSlave) {
        sheet_name = sheet_name.split('@')[0].trim();
      }

      let sheet_setting = settings[sheet_name];

      let parsed_sheet = parseSheet(sheet, sheet_setting);

      parsed_workbook[sheet_name] = parsed_sheet;

    });

    for (let name in settings) {
      if (settings[name].type === SheetType.MASTER) {

        let master_sheet = parsed_workbook[name];

        settings[name].slaves.forEach(slave_name => {

          let slave_setting = settings[slave_name];
          let slave_sheet = parsed_workbook[slave_name];

          let key_cell = _.find(slave_setting.head, item => {
            return item.type === DataType.ID || item.type === DataType.IDS;
          });

          //slave 表中所有数据
          slave_sheet.forEach(row => {
            let id = row[key_cell.name];
            delete row[key_cell.name];
            if (key_cell.type === DataType.IDS) { //array
              master_sheet[id][slave_name] = master_sheet[id][slave_name] || [];
              master_sheet[id][slave_name].push(row);
            } else { //hash
              master_sheet[id][slave_name] = row;
            }
          });

          delete parsed_workbook[slave_name];
        });
      }
    }

    return parsed_workbook;
  }
};