xlsx2json document
Let excel express complex JSON format and export excel to json.


### install
`$npm install xlsx-to-json-plus`

### Quick Start
only support .xlsx formate.

```javascript

var magic = requiplusxlsx-to-json-plus');

["./example/excel/heroes.xlsx"].forEach(function (element) {

  magic.toJson(
    path.join(__dirname, element),  //source excell file.
    path.join(__dirname, "./json"), //destination json directory.
    2,  //excell heaplusne number,first line maybe commment.
    "," //array separator.
  );

});
```

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
* date:`column_name#date`.formate:`YYYY/M/D H:m:s` or `YYYY/M/D` or `YYYY-M-D H:m:s` or `YYYY-M-D`.（attention：column type must be text，date type will cause some error for now）.
* basic type (string,number,bool):we can also leave it blank(auto type aware).
* number/boolean/string array：`column_name#[]`
* object：`column_name#{}`
* object array：`column_name#[{}]
* Date type formate:`2008-12-05 16:03:00` or `2008-18-15`
* id `column_name#id`, use to generate an object type json file, the id column would become the keys of the json object, only one id column is allowed in one sheet, see more usage in test/heroes.xlsx
* id[] `column_name#id[]`, force the value to be object array, see more usage in test/stages.xlsx

### how to use external key
* you can use the external key feature to organize more complicated data.
* create a new sheet, name it with a prefix and a '@' with an exist sheet name, it is done:)
* you may use the id to relevance the data between the two sheets, see more info in test/heroes.xlsx


### The Last
* Use `,` to separate array values by default(we can set it in config file).
* Use `;` to separate object properties。
* Only support for .xlsx format for now(do not support .xls format).
