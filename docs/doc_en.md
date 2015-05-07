xlsx2json document
=========
> Let excel express complex JSON format and export excel to json.
> Can be used on windows & *nix.

### Quick Start
* config `nodejs` environment.
* setup config files `config.json`

```json
{
    "xlsx": {
       "head": 1,	// head of the excel(set "head":2 if first line is commnet and second line is real head).
        "src": "./excel/**/[^~$]*.xlsx", 	// .xlsx files that going to be exported. glob style.
        "dest": "./json",    // directory of exported .json files.
        "arraySeparator":"," // separtor of array.
    }
}
```

* Exceute `export.bat` then `./excel/*.xlsx` files will be exported to `./json` directory.

#### Example (excel .xlsx file)
| id   | weapon  | flag   | nums#[] | words#[]   | bools#[]   | objs#{}      | obj_arr#[{}]          |
| ---- | --------| ------ | ------- | ---------- | ---------- | ------------ | --------------------- |
| 123  | shield	 | true   | 1,2     | hello,world| true,true  | a:123;b:45   | a:1;b:"hi",a:2;b:"hei"|
| 456  | sword   | false  | 3,5,8   | oh god     | false,true | a:11;b:22    | a:1;b:"hello"		 |

Result：

```json
[{
    "id": 123,
    "weapon": "shield",
    "flag": true,
    "nums": [1, 2],
    "words": ["hello", "world"],
    "bools": [true, true],
    "objs": {
        "a": 123,
        "b": 45
    },
    "obj_arr": [
      {"a": 1,"b": "hi"},
      {"a": 2,"b": "hei"}
    ]
}, {
    "id": 456,
    "weapon": "sword",
    "flag": false,
    "nums": [3, 5, 8],
    "words": ["oh god"],
    "bools": [false, true],
    "objs": {
        "a": 11,
        "b": 22
    },
    "obj_arr": [
      {"a": 1,"b": "hello"}
    ]
}]
```

### Supported Type
* number
* boolean
* string
* date
* object
* number-array
* boolean-array
* string-array
* object-array

### Excel Head Line Rule(use `#` to separate column name and column data type)
* string：`column_name#string`
* number：`column_name#number`
* bool：`column_name#bool`
* date:`column_name#date`.formate:`YYYY/M/D H:m:s` or `YYYY/M/D` or `YYYY-M-D H:m:s` or `YYYY-M-D`.（attention：xlsx column type must be text，if date type will cause some error for now）.
* basic type (string,number,bool):we can also leave it blank(automake type aware).
* number/boolean/string array：`column_name#[]`
* object：`column_name#{}`
* object array：`column_name#[{}]
* Date type formate:`2008-12-05 16:03:00` or `2008-18-15`

### Thanks
Inspiring by a clojure project [excel-to-json ](https://github.com/mhaemmerle/excel-to-json)。

### The Last
* Type `node index.js -h` int cmd window to show help;
* Use `,` to separate array values by default(we can set it in config file).
* Use `;` to separate object properties。
* Only support for .xlsx format for now(do not support .xls format).
* In order to make it portable, I put .\bin\node.exe(used by export.bat) into project.