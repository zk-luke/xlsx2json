@echo off
title [json导入mongo]
echo 按任意键确认开始将json导入mongo,若取消请直接关闭窗口。
@pause > nul
node index.js --import
@pause