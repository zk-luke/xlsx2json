@echo off
title [convert excel to json]
echo press any button to start.
@pause > nul
echo start converting ....
cd %~dp0
node index.js --export
echo convert over!
@pause