### xlsx2json ([English Document](./docs/doc_en.md))
> 让excel支持复杂的json格式, 将xlsx文件转成json。

### 更新日志
> 本次大版本更新为不兼容更新，不兼容部分涉及对象列、对象数组列和数组列。
* 2018-6-5 v1.0.0
  * feature: sheet名字以`!`开头则不导出此表。
  * feature: 列名字以`!`开头则不导出此列。
  * feature: 对象列支持对象嵌套，写法和JS中对象写法一致。
  * modify: 去掉对象数组列类型，统一为数组。
  * feature: 数组可嵌套数组和对象，和JS中数组写法一致(除了最外层不需要写中括号[])。
  * fix: 去掉自定义数组分隔符。
  * fix: excel单元格空的时候json中无此列的key。
  * modify: 去掉无用依赖，更新依赖库，精简代码。

### 分支

* `master`为主分支,此分支用于发布版本，包含当前稳定代码，不要往主分支直接提交代码。
* `dev`为开发分支,新功能bug修复等提交到此分支，待稳定后合并到`master`分支。
* 如需当做npm模块引用请切换到`npm`分支(尚有功能未合并，暂时不可用)。

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

#### 示例1(参考./excel/test.xlsx)
![excel](./docs/image/excel-data.png)

输出如下(因为设置了id列，输出hash map格式，如果无id列则输出数组格式)：

```json
{
  "1111": {
    "id": 1111,
    "name": "风暴之灵",
    "slogen": ["风暴之灵已走远","在这场风暴里没有安全的港湾啊，昆卡！"],
    "skill": {
      "R": {
        "name": "残影",
        "冷却时间": [3.5,3.5,3.5,3.5],
        "作用范围": 260,
        "主动技能": true,
        "levels": [
          {"level": 1,"damage": 140,"mana": 70},
          {"level": 2,"damage": 180,"mana": 80}
        ]
      },
      "E": {
        "name": "电子漩涡",
        "冷却时间": [21,20,19,18],
        "主动技能": true,
        "levels": [
          {"level": 1,"time": 1,"cost": 100,"distance": 100},
          {"level": 2,"time": 1.5,"cost": 110,"distance": 150}
        ]
      }
    }
  },
  "1112": {
    "id": 1112,
    "name": "幽鬼",
    "slogen": null,
    "skill": null
  }
}
```

### 支持以下数据类型

* number 数字类型。
* boolean  布尔。
* string 字符串。
* date 日期类型。
* object 对象，同JS对象一致。
* array  数组，同JS数组一致。
* id 主键类型(当表中有id类型时，json会以hash map格式输出，否则以数组格式输出)。

### 表头规则

* 基本数据类型(string,number,bool)时候，一般不需要设置会自动判断，但是也可以明确声明数据类型。
* 字符串类型：命名形式 `列名#string` 。
* 数字类型：命名形式 `列名#number` 。
* 日期类型：`列名#date` 。日期格式要符合标准日期格式。比如`YYYY/M/D H:m:s` or `YYYY/M/D` 等等。
* 布尔类型：命名形式 `列名#bool` 。
* 数组：命名形式 `列名#[]`。
* 对象：命名形式 `列名#{}` 。
* 主键：命名形式`列名#id` ,设置此将会输出为hash map 格式，否则为数组格式。
* 列名字以`!`开头则不导出此列。

### 注意事项

* 解析excel字符串的时候用到`eval()`函数，如果生产环境下excel数据来自用户输入，会有注入风险请慎用。
* 关键符号都是英文半角符号，和JS中一致。
* 对象写法同JS中对象写法一致。
* 数组写法同JS中数组写法一致(只是最外层数组不需要写中括号[])。
* sheet名字以`！`开头则不导出此表。

### TODO

- [ ] 外键支持。
- [ ] 将主分支的代码合并到npm分支。

### 补充

* 可在windows/mac/linux下运行。
* 项目地址 [xlsx2json master](https://github.com/koalaylj/xlsx2json)
* 如有问题可以到QQ群内讨论：223460081
* 招募协作开发者，有时间帮助一起维护下这个项目，可以发issue或者到qq群里把你github邮箱告诉我。
