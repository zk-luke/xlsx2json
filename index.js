'use strict'

var xlsx = require('./lib/xlsx-to-json.js');
var path = require('path');
var shell = require('child_process');
var fs = require('fs');

var config = {
    "head": 2, //表头所在的行
    "src": path.join(__dirname, "excel"), //excel目录。
    "dest": path.join(__dirname, "json") //目标目录
};

/**
 * 数据库连接配置
 * 如果没有用授权方式启动mongo 不需要填 user & pwd 并且命令行参数要加 --noauth
   比如：node index.js --import --noauth
 */
var db = {
    "host": "192.168.0.156",
    "database": "test",
    "user": "princess",
    "pwd": "princess"
};

var commands = {
    "--help": {
        "alias": ["-h", "/?"],
        "desc": "command line user manual.",
        "action": showHelp
    },
    "--export": {
        "alias": ["-e"],
        "desc": "export excel to json.",
        "action": exportJson,
        "default": true
    },
    "--import": {
        "alias": ["-i"],
        "desc": "import json to mongo.",
        "action": importMongo
    },
    "--noauth": {
        "alias": ["-na"],
        "desc": "import json to mongo without auth(do not need username and password)."
    }
};

var keys = []; //缓存命令的key ("--help"...)
var alias_map = {}; // 别名 -> key 映射
var commandLine = [];

// process.on('uncaughtException', function(err) {
//     console.log('shit 出错啦: ' + err);
// });

/**
 * 初始化
 */
(function() {
    keys = Object.keys(commands);

    for (var key in commands) {
        var alias_array = commands[key]["alias"];

        alias_array.forEach(function(element, index, array) {
            alias_map[element] = key;
        });
    };

    parseCommandLine(process.argv);

})();


/**
 * 导出json
 */
function exportJson() {
    xlsx.parse(config);
};

/**
 * 导入数据库
 */
function importMongo() {

    var uri = 'mongoimport -h ${host} -u ${user} -p ${pwd} -d ${database} -c ${collection} ${json} --jsonArray';

    //目录下两个文件 buff.json 和 palyer.json
    fs.readdir(config.dest, function(err, files) {
        if (err) {
            console.error("error", err);
            throw err;
        };

        var filtered = files.filter(function(element) {
            return path.extname(element) === '.json';
        });

        uri = uri.replace("${host}", db.host).replace("${database}", db.database);

        if (commandLine.indexOf('--noauth') === -1) {
            uri = uri.replace("${user}", db.user).replace("${pwd}", db.pwd);
        } else {
            uri = uri.replace("-u ${user} -p ${pwd}", "");
        };

        filtered.forEach(function(element, index, array) {
            var json = path.resolve(config.dest, element);
            var collection = path.basename(element, ".json");
            var newUri = uri.replace("${collection}", collection).replace("${json}", json);

            var child = shell.exec(newUri,
                function(error, stdout, stderr) {

                    console.log(collection + "\t--> " + db.host + "/" + db.database + ":" + collection);
                    console.log('state: ' + stdout);

                    if (stderr) {
                        console.error('stderr: ' + stderr);
                    };

                    if (error) {
                        console.error('exec error: ' + error);
                    }
                });
        });
    });
};

/**
 * 显示帮助
 */
function showHelp() {
    var usage = "usage:\n";
    for (var p in commands) {
        if (typeof commands[p] !== "function") {
            usage += "\t" + p + "\t" + commands[p].alias + "\t" + commands[p].desc + "\n";
        };
    };
    console.log(usage);
};


/**************************** 命令行解析 *********************************/

/**
 *  执行命令
 */
function exec(cmd_name, args) {
    if (typeof commands[cmd_name].action === "function") {
        commands[cmd_name].action(args);
    };
};

/**
 * 解析命令行参数
 */
function parseCommandLine(args) {

    if (args.length <= 2) {
        exec(defaultCommand());
    } else {
        commandLine = args.slice(2);

        //去重
        var temp = {};
        commandLine.forEach(function(element, index, array) {
            temp[element] = element;
        });

        commandLine = [];
        for (var p in temp) {
            commandLine.push(p);
        };

        commandLine.forEach(function(element, index, array) {
            if (element.indexOf('--') === -1 && element.indexOf('-') === 0) {
                commandLine[index] = alias_map[element];
                // element = alias_map[element];
            };
        });

        commandLine.forEach(function(element, index, array) {
            exec(commandLine[index]);
        });
    };
};

/**
 * 当命令行未提供参数是，默认执行的命令
 */
function defaultCommand() {
    if (keys.length <= 0) {
        throw new Error("Error: there is no command at all!");
    };

    for (var p in commands) {
        if (commands[p]["default"]) {
            return p;
        };
    };

    if (keys["--help"]) {
        return "--help";
    } else {
        return keys[0];
    };
};

/*************************************************************************/