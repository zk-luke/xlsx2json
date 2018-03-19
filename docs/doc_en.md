xlsx2json document
> Let excel express complex JSON format and export excel to json.
> Can be used on mac/linux/windows platform.

### Usage

* To clone and run this repository you'll need [Git](https://git-scm.com) and [Node.js](https://nodejs.org/en/download/) (which comes with [npm](http://npmjs.com)) installed on your computer. From your command line:
```bash
# Clone this repository
git clone https://github.com/koalaylj/xlsx2json.git
# Go into the repository
cd xlsx2json
# Install dependencies and run the app
npm install && npm start
```

* setup config files `config.json`
```javascript
{
    "xlsx": {
        /**
         * head line number.
         * first line maybe comment ,second line is head line.
         */
        "head": 2,

        /**
         * xlsx files ,glob styles.
         */
        "src": "./excel/**/[^~$]*.xlsx",

        /**
         * path for saving JSON files
         */
        "dest": "./json",

        /**
         * separator for array.
         */
        "arraySeparator":","
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
* date:`column_name#date`. regular date formate.such as `YYYY/M/D H:m:s` or `YYYY-M-D` etc..
* basic type (string,number,bool):we can also leave it blank(automake type aware).
* number/boolean/string array：`column_name#[]`
* object：`column_name#{}`
* object array：`column_name#[{}]`


### Thanks
Inspiring by a clojure project [excel-to-json ](https://github.com/mhaemmerle/excel-to-json)。

### The Last
* Use `;` to separate object properties。
* Only support for .xlsx format for now(do not support .xls format).
