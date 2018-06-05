### xlsx2json ([English Document](./docs/doc_en.md))
> 让excel支持表达复杂的json格式,将xlsx文件转成json。

### 日志
* 2017-10-31 v0.3.2
  * sheet名字以`#`开头则不导出此表。
  * 增加输出d.ts功能。
  * 去掉config.json中的arraySeparator配置。

### 分支

* `master`为主分支,此分支用于发布版本，包含当前稳定代码，不要往主分支直接提交代码。
* `dev`为开发分支,新功能提交到此分支，待稳定后合并到`master`分支。
* 如需当做npm模块引用请切换到`npm`分支(尚有功能未合并)。

### 使用说明
* 目前只支持.xlsx格式，不支持.xls格式。
* 本项目是基于nodejs的，所以需要先安装nodejs环境。
* 执行命令
```bash
# Clone this repository
git clone https://github.com/koalaylj/xlsx2json.git
# Go into the repository
cd xlsx2json
# Install dependencies
npm install
```

* 配置config.json
```javascript
{
    "xlsx": {
        /**
         * 表头所在的行，第一行可以是注释，第二行是表头
         */
        "head": 2,

        /**
         * xlsx文件所在的目录
         * glob配置风格
         */
        "src": "./excel/**/[^~$]*.xlsx",

        /**
         * 导出的json存放的位置
         */
        "dest": "./json"
    },

    "ts":false,//是否导出d.ts（for typescript）

    "json": {
      /**
       * 导出的json是否需要压缩
       * true:压缩，false:不压缩(便于阅读的格式)
       */
      "uglify": false
    }
}
```
* 执行`export.sh|export.bat`即可将`./excel/*.xlsx` 文件导成json并存放到 `./json` 下。json名字以excel的sheet名字命名。

* 补充(一般用不上)：
    * 执行`node index.js -h` 查看使用帮助。
    * 命令行传参方式使用：执行 node `index.js --help` 查看。

#### 示例1 test.xlsx
| id   | desc         | flag   | nums#[] | words#[]    |   map#[]/   | data#{}      | hero#[{}]                     |
| ---- | -------------| ------ | ------- | ----------- | ---------- | ------------ | --------------------------    |
| 123  | description  | true   | 1,2     | 哈哈,呵呵     | true/true  | a:123;b:45   | id:2;level:30,id:3;level:80  |
| 456  | 描述          | false  | 3,5,8   | shit,my god | false/true | a:11;b:22    | id:9;level:38,id:17;level:100 |


输出如下：

```json
[{
    "id": 123,
    "desc": "description",
    "flag": true,
    "nums": [1, 2],
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
    "nums": [3, 5, 8],
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

### 支持以下数据类型

* number 数字类型。
* boolean  布尔。
* string 字符串。
* date 日期类型。
* object 简单对象，暂时不支持对象里面有对象或数组这种。
* number-array  数字数组。
* boolean-array  布尔数组。
* string-array  字符串数组。
* object-array 对象数组。
* id 主键类型(当表中有这个类型的时候，json会以map格式输出，否则以数组格式输出)。

### 表头规则

* 基本数据类型(string,number,bool)时候，一般不需要设置会自动判断，但是也可以明确声明数据类型。
* 字符串类型：命名形式 `列名#string` 。
* 数字类型：命名形式 `列名#number` 。
* 日期类型：`列名#date` 。日期格式要符合标准日期格式。比如`YYYY/M/D H:m:s` or `YYYY/M/D` 等等。
* 布尔类型：命名形式 `列名#bool` 。
* 基本类型数组：命名形式 `列名#[]`，数组元素默认用逗号分隔(a,b,c),自定义数组元素分隔符`列名#[]/`(a/b/c)。
* 对象：命名形式 `列名#{}` 。
* 对象数组：命名形式`列名#[{}]` 。
* 主键：命名形式`列名#id` ,设置此将会输出为json map 格式。

### 规则

* 关键符号都是半角符号。
* 对象属性使用分号`;`分割。
* sheet名字以`#`开头则不导出此表。

### TODO

- [ ] 列为数组类型时候，嵌套复杂类型。
- [ ] 列为对象类型时候，嵌套复杂类型。
- [ ] 外键支持。
- [ ] 将主分支的代码合并到npm分支。

### 补充

* 可在windows/mac/linux下运行。
* 项目地址 [xlsx2json master](https://github.com/koalaylj/xlsx2json)
* 如有问题可以到QQ群内讨论：223460081
* 招募协作开发者，有时间帮助一起维护下这个项目，可以发issue或者到qq群里把你github邮箱告诉我。
