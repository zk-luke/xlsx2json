'use strict'

var xlsx = require('./lib/xlsx-to-json.js');
var path = require('path');
var shell = require('child_process');
var fs = require('fs');
var glob = require("glob");


/**
 * 命令支持的参数
 */
var commands = {
    "--help": {
        "alias": ["-h"],
        "desc": "show this help manual.",
        "action": showHelp
    },
    "--export": {
        "alias": ["-e"],
        "desc": "export excel to json. --export [files]",
        "action": exportJson,
        "default": true
    },
    "--import": {
        "alias": ["-i"],
        "desc": "import json to mongo. --import [files]",
        "action": importMongo
    },
    "--noauth": {
        "alias": ["-na"],
        "desc": "import json to mongo without auth(do not need username and password)."
    }
};


/**
 * 配置文件
 * 如果没有用授权方式启动mongo 不需要填 user & pwd 并且命令行参数要加 --noauth
   比如：node index.js --import --noauth
 */
var config = {
    "head": 2, //表头所在的行
    "export": { //excel目录和json目录
        "from": "./excel/**/[^~$]*.xlsx",
        "to": "./json"
    },
    "import": { //数据库相关信息
        // "from": "./json/**/*.json",
        "to": {
            "host": "127.0.0.1",
            "database": "test",
            "user": "princess",
            "pwd": "princess"
        }
    }
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

    // console.log("%j", parsed_cmds);

    parsed_cmds.forEach(function(element, index, array) {
        exec(element);
    });
})();


/**
 * 导出json
 * args: --export 命令行参数 excel文件列表，如果为空，则按照config.export.from导出。
 */
function exportJson(args) {

    if (typeof args === 'undefined' || args.length == 0) {
        glob(config.export.from, function(err, files) {
            if (err) {
                console.error("config.export.from match error:", err);
                throw err;
            };

            console.log(files);

            files.forEach(function(element, index, array) {
                xlsx.toJson(path.join(__dirname, element), path.join(__dirname, config.export.to), config.head);
            });

        })
    } else {
        if (args instanceof Array) {
            args.forEach(function(element, index, array) {
                xlsx.toJson(path.join(__dirname, element), path.join(__dirname, config.export.to), config.head);
            });
        };
    }
};


/**
 * 导入数据库
 */
function importMongo(args) {

    var uri = 'mongoimport -h ${host} -u ${user} -p ${pwd} -d ${database} -c ${collection} ${json} --jsonArray';

    var files_to_import = args;

    if (typeof args === 'undefined' || args.length == 0) {
        files_to_import = fs.readdirSync(config.export.to);
        files_to_import.forEach(function(element, index, array) {
            array[index] = path.join(config.export.to, element);
        });
    }

    console.log("files_to_import", files_to_import);

    var db = config.import.to;

    uri = uri.replace("${host}", db.host).replace("${database}", db.database);

    if (keys.indexOf('--noauth') === -1) {
        uri = uri.replace("${user}", db.user).replace("${pwd}", db.pwd);
    } else {
        uri = uri.replace("-u ${user} -p ${pwd}", "");
    };

    files_to_import.forEach(function(element, index, array) {
        if (path.extname(element) === '.json') {

            var temp = element.split(path.sep);
            var collection = path.basename(temp[temp.length - 1], ".json");
            var newUri = uri.replace("${collection}", collection).replace("${json}", element);

            var child = shell.exec(newUri, function(error, stdout, stderr) {

                console.log(collection + "\t--> " + db.host + "/" + db.database + ": " + collection);

                console.log('state: ' + stdout);

                if (stderr) {
                    console.error('stderr: ' + stderr);
                };

                if (error) {
                    console.error('exec error: ' + error);
                }
            });
        };
    });
};

/**
 * 显示帮助
 */
function showHelp() {
    var usage = "usage: \n";
    for (var p in commands) {
        if (typeof commands[p] !== "function") {
            usage += "\t " + p + "\t " + commands[p].alias + "\t " + commands[p].desc + "\n ";
        };
    };

    usage += "\nexamples: ";
    usage += "\n\n $node index.js--export\n\tthis will export all files configed to json.";
    usage += "\n\n $node index.js--export. / excel / foo.xlsx. / excel / bar.xlsx\n\tthis will export foo and bar xlsx files.";
    usage += "\n\n $node index.js--import\n\tthis will import all configed json files.";
    usage += "\n\n $node index.js--import. / json / foo.json. / excel / bar.json\n\tthis will export foo and bar json files.";
    usage += "\n\n $node index.js--import--noauth\n\t db do not need auth(do not need db user & pwd).";

    console.log(usage);
};


/**************************** 命令行解析 *********************************/

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
                    throw new Error("not support command:" + cli[index]);
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