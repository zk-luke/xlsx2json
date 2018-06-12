const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');
var _ = require('lodash');
const config = require('../config.json');
const types = require('./types');
const parser = require('./parser');

const DataType = types.DataType;

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
 * save workbook
 */
function serializeWorkbook(parsedWorkbook, dest) {

  for (let name in parsedWorkbook) {
    let sheet = parsedWorkbook[name];

    let dest_file = path.resolve(dest, name + ".json");
    let resultJson = JSON.stringify(sheet, null, config.json.uglify ? 0 : 2); //, null, 2

    fs.writeFile(dest_file, resultJson, err => {
      if (err) {
        console.error("error：", err);
        throw err;
      }
      console.log('exported json  -->  ', path.basename(dest_file));
    });

  }

}


/**
 * save dts
 */
function serializeDTS(dest, fileName, settings) {

  let dts = "";

  for (let name in settings) {
    dts += formatDTS(name, settings[name]);
  }

  let dest_file = path.resolve(dest, fileName + ".d.ts");
  fs.writeFile(dest_file, dts, err => {
    if (err) {
      console.error("error：", err);
      throw err;
    }
    console.log('exported t.ds  -->  ', path.basename(dest_file));
  });

}


/**
 * 
 * @param {String} name the excel file name will be use on create d.ts
 * @param {Object} head the excel head will be the javescript field
 */
function formatDTS(name, setting) {
  let className = _.capitalize(name);
  let strHead = "interface " + className + " {\r\n";
  for (let i = 0; i < setting.head.length; ++i) {
    let head = setting.head[i];
    if (head.name.startsWith('!')) {
      continue;
    }
    let typesDes = "any";
    switch (head.type) {
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
      case DataType.UNKNOWN:
        {
          typesDes = "any";
          break;
        }
      default:
        {
          typesDes = "any";
        }
    }
    strHead += "\t" + head.name + ": " + typesDes + "\r\n";
  }

  setting.slaves.forEach(slave_name => {
    strHead += "\t" + slave_name + ": " + _.capitalize(slave_name) + "\r\n";
  });

  strHead += "}\r\n";
  return strHead;
}

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
  toJson: function (src, dest) {

    if (!fs.existsSync(dest)) {
      fs.mkdirSync(dest);
    }

    let parsed_src = path.parse(src);

    let workbook = xlsx.parse(src);

    console.log("parsing excel:", parsed_src.base);

    let settings = parser.parseSettings(workbook);

    // let parsed_workbook = parseWorkbook(workbook, dest, headIndex, path.join(dest, parsed_src.name));
    let parsed_workbook = parser.parseWorkbook(workbook, settings);

    serializeWorkbook(parsed_workbook, dest);

    if (config.ts) {
      serializeDTS(dest, parsed_src.name, settings);
    }
  }
};