@echo off
title [excell导出json]
echo 按任意键确认开始将excel导出json,若取消请直接关闭窗口。
@pause > nul
.\bin\node index.js --export
@pause