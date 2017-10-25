const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');
const moment = require('moment');

let arraySeparator;

module.exports = {

  /**
   * convert xlsx file to json and save it to file system.
   * @param  {String} src_excel_file path of .xlsx files.
   * @param  {String} dest_dir       directory for exported json files.
   * @param  {Number} headIndex      index of head line.
   * @param  {String} separator      array separator.
   */
  toJson: function(src_excel_file, dest_dir, headIndex, separator) {

    arraySeparator = separator;

    if (!fs.existsSync(dest_dir)) {
      fs.mkdirSync(dest_dir);
    }

    console.log("parsing excel:", src_excel_file);

    let workbook = xlsx.parse(src_excel_file);
    parseWorkbook(workbook, dest_dir, headIndex);
  }
};

/**
 * convert worksheet in workbook and save to file for each.
 * @param       {[Object]} workbook json object of excel's workbook.
 * @param       {[String]} dest     directory for exported json files.
 * @param       {[Number]} headIndex     index of head line.
 */
function parseWorkbook(workbook, dest, headIndex) {

  workbook.forEach(sheet => {

    // ignore sheet with external keys only, or start with a '#'
    if (sheet.name.indexOf('@') === -1 && sheet.name[0] !== "#") {
      let parsedSheet = parseSheet(sheet, headIndex);

      let dest_file = path.resolve(dest, sheet.name + ".json");
      fs.writeFile(dest_file, JSON.stringify(parsedSheet, null, 2), err => {
        if (err) {
          console.error("error：", err);
          throw err;
        }
        console.log('exported successfully  -->  ', path.basename(dest_file));
      });
    }
  });

  return;

  //TODO
  for (let i_sheet = 0; i_sheet < workbook.worksheets.length; i_sheet++) {
    let sheet = workbook.worksheets[i_sheet];
    console.log("sheet:" + sheet.name + " cols:" + sheet.maxCol, "rows:" + sheet.maxRow);

    // process the sheet without external keys only, or start with a '#'
    if (sheet.name.indexOf('@') === -1 && sheet.name[0] !== "#") {
      let output = parseSheet(sheet, head);

      // scan rest sheets for external keys
      // for (let j = i_sheet; j < excel.worksheets.length; j++) {
      //   let rest_sheet = excel.worksheets[j];
      //   let rest_sheet_name = rest_sheet.name;
      //   if ((rest_sheet_name + '').indexOf('@') != -1) {
      //     let temp = rest_sheet_name.split('@');
      //     let external_key = temp[0];
      //     let sheet_name = temp[1];
      //     let external_values = parseSheet(rest_sheet, head);
      //     if (sheet.name === sheet_name) {
      //       console.log("find external_key:" + external_key, "in sheet:" + sheet_name);
      //       for (let k in output) {
      //         //console.log("k:" + k, "output[k]:" + output[k]);
      //         if (output[k] && external_values[k]) {
      //           output[k][external_key] = external_values[k];
      //         }
      //       }
      //     }
      //   }
      // }

      output = JSON.stringify(output, null, 2);

      let dest_file = path.resolve(dest, sheet.name + ".json");
      fs.writeFile(dest_file, output, err => {
        if (err) {
          console.log("error：", err);
          throw err;
        }
        console.log('exported successfully  -->  ', path.basename(dest_file));
      });
    }
  }
}

let OUTPUT_ARRAY = 0;
let OUTPUT_OBJ_VALUE = 1;
let OUTPUT_OBJ_ARRAY = 2;

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

    let objOutput = OUTPUT_ARRAY;

    let result = objOutput ? {} : [];

    for (let i_row = headIndex + 1; i_row < sheet.data.length; i_row++) {

      let row = sheet.data[i_row];

      // console.log("\t\t row", i_row, row);

      // let id;

      let parsedRow = parseRow(row, i_row, head);
      result.push(parsedRow);

      // console.log("objOutput", objOutput);

      //TODO
      // if (objOutput === OUTPUT_OBJ_VALUE) {
      //   output[id] = json_obj;
      // } else if (objOutput === OUTPUT_OBJ_ARRAY) {
      //   output[id].push(json_obj);
      // } else if (objOutput === OUTPUT_ARRAY) {
      //   output.push(json_obj);
      // }
    }
    //console.log("output******",output);
    return result;
  }
}

function parseHead(sheet, headIndex) {

  let headRow = sheet.data[headIndex];

  // console.log("\t\t parsing head", headIndex, headRow);

  let head = {
    names: [],
    types: [],
    index: headIndex
  };

  headRow.forEach(cell => {
    let type = 'basic';
    let name = cell;

    if ((cell + '').indexOf('#') !== -1) {
      let pair = cell.split('#');
      name = pair[0];
      type = pair[1].toLowerCase().trim();

      //TODO if there exists an id type, change the whole output to a json object
      // if (type === 'id') {
      //   objOutput = OUTPUT_OBJ_VALUE;
      // } else if (type === 'id[]') {
      //   objOutput = OUTPUT_OBJ_ARRAY;
      // }
    }

    head.types.push(type);
    head.names.push(name);
  });

  return head;
}

function parseRow(row, rowIndex, head) {

  let result = {};
  let id;

  row.forEach((cell, index) => {

    if (cell) {

      let name = head.names[index];
      let type = head.types[index];

      switch (type) {
        case 'id': // id
          id = cell;
          break;
        case 'id[]': //TODO id[]
          // if (cell) {
          //   id = cell;
          //   if (!output[id]) {
          //     output[id] = [];
          //   }
          // }
          break;
        case 'basic': // number string boolean
          result[name] = cell;
          break;
        case 'date':
          if (isDateType(cell)) {
            result[name] = cell;
          } else {
            //xlsx's bug!!!
            result[name] = numdate(cell);
          }
          break;
        case 'string':
          result[name] = cell;
          break;
        case 'number':
          //+xxx.toString() '+' means convert it to number
          let isNumber = !isNaN(+cell.toString());
          if (isNumber) {
            result[name] = Number(cell);
          } else {
            console.warn("type error at [" + rowIndex + "," + index + "]," + cell + " is not a number");
          }
          break;
        case 'bool':
          result[name] = toBoolean(cell.toString());
          break;
        case '{}': //support {number boolean string date} property type
          result[name] = array2object(cell.split(';'));
          break;
        case '[]': //[number] [boolean] [string]  todo:support [date] type
          result[name] = parseBasicArrayField(cell);
          break;
        case '[{}]':
          result[name] = parseObjectArrayField(cell);
          break;
        default:
          console.log('unrecognized type', cell, typeof(cell));
          break;
      }
    }
  });
  return result;
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
      let kv = e.trim().split(':');
      if (isNumber(kv[1])) {
        kv[1] = Number(kv[1]);
      } else if (isBoolean(kv[1])) {
        kv[1] = toBoolean(kv[1]);
      }
      result[kv[0]] = kv[1];
    }
  });
  return result;
}

/**
 * parse object
 */
// function parseObject(field, key, data) {
//   field[key] = array2object(data.split(';'));
// }

/**
 * parse simple array.
 */
function parseBasicArrayField(array) {
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

  if (value) {
    if (typeof value === 'number') {
      return true;
    }
    return !isNaN(+value.toString());
  }

  return false;
}

/**
 * boolean type check.
 */
function isBoolean(value) {

  if (typeof(value) == "undefined") {
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
function isDateType(value) {
  if (isNumber(value)) {
    return false;
  }
  if (value) {
    return moment(new Date(value), "YYYY-M-D", true).isValid() || moment(value, "YYYY-M-D H:m:s", true).isValid() || moment(value, "YYYY/M/D H:m:s", true).isValid() || moment(value, "YYYY/M/D", true).isValid();
  }
  return false;
}

//fuck xlsx's bug
var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
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
