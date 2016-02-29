[![Build Status](https://travis-ci.org/koalaylj/xlsx2json.svg?branch=master)](https://travis-ci.org/koalaylj/xlsx2json)
xlsx2json
=========
[English Document](./docs/doc_en.md)


### 作用
让excel表达复杂的json格式,将xlsx文件转成json。

### npm相关
* 如需当做npm模块引用请切换到`npm`分支。

### 感谢
某些想法也是借鉴了一个clojure的excel转json的开源项目 [excel-to-json ](https://github.com/mhaemmerle/excel-to-json)。


### 使用说明
> 目前只支持.xlsx格式，不支持.xls格式。

* 首先设置好node环境变量。
* 配置config.json
```json
{
    "xlsx": {
        "head": 2,			// 表头所在的行，第一行可以是注释，第二行是表头。
        "src": "./excel/**/[^~$]*.xlsx", // xlsx文件 glob配置风格
        "dest": "./json",	 //  导出的json存放的位置
        "arraySeparator":","  // 数组的分隔符
    }
}
```
* 执行`export.sh/export.bat`即可将`./excel/*.xlsx` 文件导成json并存放到 `./json` 下。json名字以excel的sheet名字命名。

* 补充(一般用不上)：
    * 执行`node index.js -h` 查看使用帮助。
    * 命令行传参方式使用：执行 node `index.js --help` 查看。

#### 示例1 test.xlsx
![test.xlsx](http://img3.douban.com/view/photo/photo/public/p2180848214.jpg)

输出如下：

```json
[{
    "id": 123,
    "desc": "description",
    "flag": true,
    "otherid": [1, 2],
    "words": ["哈哈", "呵呵"],
    "map": [true, true],
    "data": {
        "a": 123,
        "b": 45
    },
    "hero": [
      {"id": 2,"level": 30},
      {"id": 3,"level": 80}
    ]
}, {
    "id": 456,
    "desc": "描述",
    "flag": false,
    "otherid": [3, 5, 8],
    "words": ["shit", "my god"],
    "map": [false, true],
    "data": {
        "a": 11,
        "b": 22
    },
    "hero": [
      {"id": 9, "level": 38 },
      {"id": 17,"level": 100}
    ]
}]
```

## 支持以下数据类型
* number 数字类型
* boolean  布尔
* string 字符串
* date 日期类型
* object 对象，复杂的嵌套可以通过外键来实现，见“外键类型的sheet关联”
* number-array  数字数组
* boolean-array  布尔数组
* string-array  字符串数组
* object-array 对象数组，复杂的嵌套可以通过外键来实现，见“外键类型的sheet关联”

## 表头规则
* 基本数据类型(string,number,bool)时候，一般不需要设置会自动判断，但是也可以明确声明数据类型。
* 字符串类型：此列表头的命名形式 `列名#string` 。
* 数字类型：此列表头的命名形式 `列名#number` 。
* 日期类型：`列名#date` 。格式`YYYY/M/D H:m:s` or `YYYY/M/D` or `YYYY-M-D H:m:s` or `YYYY-M-D`。（==注意：目前xlsx文件里面列必须设置为文本类型，如果是日期类型的话，会导致底层插件解析出来错误格式的时间==）.
* 布尔类型：此列表头的命名形式 `列名#bool` 。
* 基本类型数组：此列表头的命名形式 `列名#[]` 。
* 对象：此列表头的命名形式 `列名#{}` 。
* 对象数组：此列表头的命名形式`列名#[{}]` 。
* id：此列表头的命名形式`列名#id`，用来生成对象格式的输出，以该列字段作为key，一个sheet中不能存在多个id类型的列，否则会被覆盖，相关用例请查看test/heroes.xlsx
* id[]：此列表头的命名形式`列名#id[]`，用来约束输出的值为对象数组，相关用例请查看test/stages.xlsx

## 数据规则
* 关键符号都是半角符号。
* 数组使用逗号`,`分割。
* 对象属性使用分号`;`分割。
* 列格式如果是日期，导出来的是格林尼治时间不是当时时区的时间，列设置成字符串可解决此问题。

## 外键类型的sheet关联
* sheet名称必须为【列名@sheet名称】，例如存在一个名称为a的sheet，会导出一个a.json，可以使用一个名称为b@a的sheet为这个json添加一个b的属性
* 外键类型的sheet（sub sheet）顺序上必须位于被关联的sheet（master sheet）之后
* master sheet的输出类型如果为对象，则sub sheet必须也存在master sheet同列名并且类型为id的列作为关联关系；master sheet的输出类型如果为数组，则sub sheet按照数组下标（行数）顺序关联
* 相关用例请查看test/heroes.xlsx

## 原理说明
* 依赖 `node-xlsx` 这个项目解析xlsx文件。
* xlsx就是个zip文件，解压出来都是xml。有一个xml存的string，有相应个xml存的sheet。通过解析xml解析出excel数据(json格式)，这个就是`node-xlsx` 做的工作。
* 本项目只需利用 `node-xlsx` 解析xlsx文件，然后拼装自己的json数据格式。

## 补充
* windows/mac/linux都支持。
* 项目地址 [xlsx2json master](https://github.com/koalaylj/xlsx2json)
* 如有问题可以到QQ群内讨论：223460081
* 项目中的某些工具函数测试用例请参见我的gist js:validate & js:convert。
