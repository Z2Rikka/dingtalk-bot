@echo off
chcp 65001 >nul
title 钉钉文档收集机器人

echo ========================================
echo   钉钉文档收集机器人
echo ========================================
echo.

cd /d "%~dp0"

echo [1/3] 检查Node.js...
node --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Node.js 未安装，请先安装Node.js
    pause
    exit /b 1
)
echo ✅ Node.js 已安装: 
node --version

echo.
echo [2/3] 检查依赖...
if not exist node_modules (
    echo 📦 正在安装依赖...
    call npm install
    if errorlevel 1 (
        echo ❌ 依赖安装失败
        pause
        exit /b 1
    )
    echo ✅ 依赖安装完成
) else (
    echo ✅ 依赖已安装
)

echo.
echo [3/3] 检查配置...
node -e "try { require('./config.js'); console.log('✅ 配置文件正常'); } catch(e) { console.log('⚠️ 配置文件有问题: ' + e.message); process.exit(1); }"
if errorlevel 1 (
    echo.
    echo ⚠️ 请先配置 config.js 文件
    echo.
    echo 请编辑 config.js 填写以下内容：
    echo   - bot: appKey, appSecret, agentId
    echo   - storage: baseDir 存储目录
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo ✅ 启动机器人
echo ========================================
echo.

start "钉钉文档收集机器人" cmd /k "cd /d %~dp0 && node bot.js"

echo.
echo ✅ 启动成功！
echo.
echo 回调地址：http://你的服务器IP:3000/webhook
echo.
echo 常用命令：
echo   /帮助   - 显示帮助
echo   /状态   - 收集统计
echo   /列表   - 最近文件
echo   /目录   - 存储目录
echo.

pause
