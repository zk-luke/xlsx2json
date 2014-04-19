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

var keys = []; //缓存commands的key ("--help"...)
var alias_map = {}; // alias_name -> name 映射
var parsed_cmds = []; //解析命令行出来的命令

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

    parsed_cmds = parseCommandLine(process.argv);

    parsed_cmds.forEach(function(element, index, array) {
        exec(element);
    });
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
function importMongo(args) {

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

        if (parsed_cmds.indexOf('--noauth') === -1) {
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
 * 命令支持的参数
 */
var commands = {
    "--help": {
        "alias": ["-h"],
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


/**
 * 执行一条命令
 */
function exec(cmd) {
    if (typeof cmd.action === "function") {
        cmd.action(cmd.args);
    };
};


/**
 * 解析命令行参数
 */
function parseCommandLine(args) {

    var parsed_cmds = [];

    if (args.length <= 2) {
        parsed_cmds.push(defaultCommand());
    } else {

        var cli = args.slice(2);

        var pos = 0;
        var cmd;

        cli.forEach(function(element, index, array) {

            //将别名替换为正式命令名字
            if (element.indexOf('--') === -1 && element.indexOf('-') === 0) {
                cli[index] = alias_map[element];
            }

            //解析命令和相应的参数
            if (cli[index].indexOf('--') === -1) {
                cmd.args.push(cli[index]);
            } else {

                if (keys[cli[index]] == "undefined") {
                    throw new Error("not support command: " + cli[index]);
                };

                pos = index;
                cmd = commands[cli[index]];
                if (typeof cmd.args == 'undefined') {
                    cmd.args = [];
                };
                parsed_cmds.push(cmd);
            }
        });
    };

    return parsed_cmds;
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
            return commands[p];
        };
    };

    if (keys["--help"]) {
        return commands["--help"];
    } else {
        return commands[keys[0]];
    };
};

/*************************************************************************/