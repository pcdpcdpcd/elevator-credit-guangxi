@echo off
chcp 65001 >nul
title 电梯维保单位信用评分工具
echo 正在启动，请勿关闭本窗口...
echo 启动后请浏览器访问 http://localhost:3000
node server.js
pause
