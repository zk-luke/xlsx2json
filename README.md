xlsx2json
=========

可以让excel表达复杂的json格式，将excel转成json存入mongo数据库。

## 产生背景

开发游戏的时候，策划用的是excel，而我们的数据库是mongo。
我们的工作流是策划写excel，然后导出json，然后倒入mongo数据库。

因为excell是二维的，无法表达mongo document里面数组和嵌套对象结构。
导致服务器端只能按照excel结构来设计mongo数据库，
只能和mysql等关系型数据库以同样的方式的使用和设计mongo，
这样会使设计mongo数据库大大受限并无法使用mongo的某些特性。

有一个clojure项目 [excel-to-json ](https://github.com/mhaemmerle/excel-to-json) 可以完成这个功能。
但是不懂clojure表示压力很大而且有些功能不符合我们的需求，so就搞了这个项目。
某些想法也是借鉴了这个项目，在此表示感谢。

## 说明
* 项目只依赖 `node-xlsx` 项目，解析xlsx文件的。
* 其实xlsx就是个zip文件，解压出来都是xml。
  有一个xml存的string，有相应个xml存的sheet。
  只要解析xml就能解析出excel数据了，这个就是`node-xlsx` 项目做的工作。
* 只需利用 `node-xlsx` 解析xslx文件，拼装json数据。
* 项目中的某些工具函数测试用例请参见我的gist js:validate & js:convert。


## 支持以下数据类型
* number 数字类型
* boolean  布尔
* string 字符串
* object 对象 {a:1,b:false} // 对象内暂时不能有数组，也不能对象嵌套对象，此功能正在加入。
* number-array  数字数组
* boolean-array  布尔数组
* string-array  字符串数组
* object-array 对象数组


## 表头规则
* 列是基本数据类型(string,number,boolean)时候，不需要特殊设置。
* 列是基本类型数组的话 列命名形式  colume#[] ，后面要加#[] 表达这列是基本类型数组。
* 列是对象的话，命名 colume#{} ,后面要加 #{} 表达这列是对象。
* 列是对象数组话，命名 colume#[{}] ,后面要加 #[{}] 表达这列是对象数组。


## 数据规则
* 关键符号都是半角符号。
* 数组使用逗号分割。

### example  test.xlsx

| id    | desc        | flag   | otherid#[]  | words#[]     | map#[]     | data#{}      | hero#[{}]                     |
| ----- | ----------- | ------ | ----------- | ------------ | ---------- | ------------ | ----------------------------- |
| 123   | description | true   | 1,2         | 哈哈,呵呵    | true,true  | a:123;b:45   | id:2;level:30,id:3;level:80   |
| 456   | 描述        | false  | 3,5,8       | shit,my god  | false,true | a:11;b:22    | id:9;level:38,id:17;level:100 |

输出如下

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
    "map": [false, false],
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

## 配置文件 config.json

```json
{
    "head": 2,  //表头从第几行开始,第一行可能是注释。
    "debug": true, //暂时没用
    "src": ["./**/*.xlsx"], //excel文件列表 暂时没用 后续会加上
    "dest": "json" //输出目录，暂时没用 后续会加上
}
```

## 项目当前状态
* 进行中的功能：增加嵌套对象和嵌套数组支持。
* 完善配置文件功能。
* 将json导入mongo中。完善工作流。
