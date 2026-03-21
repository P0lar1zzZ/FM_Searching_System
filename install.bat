@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cls

echo ========================================
echo   防火门一键搜取 - 环境安装脚本
echo ========================================
echo.

REM ===== 第一步：检查 Python 安装 =====
echo [1/4] 正在检查 Python 安装...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ⚠️  未检测到 Python！
    echo 正在为你寻找最快的 Python 下载源...
    echo.
    
    REM 定义国内镜像源
    set "MIRROR1=https://npmmirror.com/mirrors/python"
    set "MIRROR2=https://mirrors.aliyun.com/python"
    set "MIRROR3=https://mirrors.tsinghua.edu.cn/python-releases"
    set "MIRROR_OFFICIAL=https://www.python.org/downloads/windows/"
    
    set "PYTHON_FOUND=0"
    set "FASTEST_MIRROR="
    
    REM 测试镜像可用性
    echo 正在测试镜像源...
    
    REM 测试 npmmirror（通常最快）
    echo   • 测试 npmmirror 源...
    curl -I -s -m 3 "%MIRROR1%" >nul 2>&1
    if errorlevel 0 (
        set "FASTEST_MIRROR=%MIRROR1%"
        set "PYTHON_FOUND=1"
        echo     ✓ 可用！（npmmirror CDN）
    ) else (
        echo     ✗ 超时
    )
    
    REM 如果 npmmirror 失败，测试阿里云
    if !PYTHON_FOUND! equ 0 (
        echo   • 测试阿里云镜像源...
        curl -I -s -m 3 "%MIRROR2%" >nul 2>&1
        if errorlevel 0 (
            set "FASTEST_MIRROR=%MIRROR2%"
            set "PYTHON_FOUND=1"
            echo     ✓ 可用！（阿里云）
        ) else (
            echo     ✗ 超时
        )
    )
    
    REM 如果前两个都失败，测试清华源
    if !PYTHON_FOUND! equ 0 (
        echo   • 测试清华大学镜像源...
        curl -I -s -m 3 "%MIRROR3%" >nul 2>&1
        if errorlevel 0 (
            set "FASTEST_MIRROR=%MIRROR3%"
            set "PYTHON_FOUND=1"
            echo     ✓ 可用！（清华大学）
        ) else (
            echo     ✗ 超时
        )
    )
    
    echo.
    
    if !PYTHON_FOUND! equ 1 (
        echo ✓ 找到最快镜像源，正在打开下载页面...
        echo 📍 推荐版本：Python 3.9 或 3.10
        echo 💡 重要提示：安装时务必勾选 "Add Python to PATH"
        echo.
        timeout /t 3 >nul
        start !FASTEST_MIRROR!
    ) else (
        echo ⚠️  所有镜像源均无法连接，打开官方下载页面...
        echo 📍 访问 Python 官网获取最新版本
        echo.
        timeout /t 2 >nul
        start %MIRROR_OFFICIAL%
    )
    
    echo.
    echo 请先完成 Python 安装，记得勾选 "Add Python to PATH" ✓
    echo 安装完成后，请重新运行此脚本。
    echo.
    pause
    exit /b 1
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
    echo ✓ 检测到: !PYTHON_VERSION!
)
echo.

REM ===== 第二步：一键静默安装依赖 =====
echo [2/4] 正在安装依赖包...
echo 使用加速镜像: 清华大学 TUNA 镜像
echo.

set "PACKAGES=pywin32 pandas openpyxl"
set "MIRROR=https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple"
set "SUCCESS_COUNT=0"
set "FAIL_COUNT=0"

for %%p in (%PACKAGES%) do (
    echo   • 安装 %%p...
    pip install -i !MIRROR! %%p -q >nul 2>&1
    if errorlevel 1 (
        echo     ✗ 安装失败
        set /a FAIL_COUNT=!FAIL_COUNT!+1
    ) else (
        echo     ✓ 安装成功
        set /a SUCCESS_COUNT=!SUCCESS_COUNT!+1
    )
)

echo.
echo 安装汇总: !SUCCESS_COUNT! 个成功, !FAIL_COUNT! 个失败
if !FAIL_COUNT! gtr 0 (
    echo.
    echo ⚠️  部分依赖安装失败！
    echo 请手动执行以下命令重试：
    echo.
    echo pip install -i !MIRROR! %PACKAGES%
    echo.
    echo 或者尝试其他镜像：
    echo   - 阿里云: https://mirrors.aliyun.com/pypi/simple/
    echo   - 腾讯云: http://mirrors.cloud.tencent.com/pypi/simple
    pause
) else (
    echo ✓ 所有依赖安装完毕！
)
echo.

REM ===== 第三步：从 GitHub 下载最新脚本 =====
echo [3/4] 从 GitHub 下载最新防火门脚本...
echo.

set "GITHUB_REPO=https://github.com/P0lar1zzZ/FM_Searching_System.git"
set "GITHUB_RAW=https://raw.githubusercontent.com/P0lar1zzZ/FM_Searching_System/main"
set "MIRROR_RAW=https://raw.fastgit.org/P0lar1zzZ/FM_Searching_System/main"

REM 尝试下载 fm_searching 脚本
echo   • 正在从 GitHub 下载最新脚本...

REM 优先使用加速镜像 (fastgit)
curl -s -o "%cd%\fm_searching_new.py" "%MIRROR_RAW%/fm_searching.py" >nul 2>&1
if errorlevel 1 (
    REM 加速镜像失败，尝试官方 GitHub
    echo   • 加速镜像超时，尝试官方源...
    curl -s -o "%cd%\fm_searching_new.py" "%GITHUB_RAW%/fm_searching.py" >nul 2>&1
)

if exist "%cd%\fm_searching_new.py" (
    REM 备份原文件
    if exist "%cd%\fm_searching.py" (
        move /y "%cd%\fm_searching.py" "%cd%\fm_searching.py.backup" >nul 2>&1
    )
    REM 使用新文件
    move /y "%cd%\fm_searching_new.py" "%cd%\fm_searching.py" >nul 2>&1
    echo     ✓ 脚本下载成功！
) else (
    echo     ⚠️  下载失败，继续使用本地脚本
    echo.
    echo     如需手动更新，请运行：
    echo     git clone %GITHUB_REPO%
    echo     或访问: %GITHUB_REPO%
)

echo   • 同时检查 generate_icon.py...
curl -s -o "%cd%\generate_icon_new.py" "%MIRROR_RAW%/generate_icon.py" >nul 2>&1
if errorlevel 1 (
    curl -s -o "%cd%\generate_icon_new.py" "%GITHUB_RAW%/generate_icon.py" >nul 2>&1
)

if exist "%cd%\generate_icon_new.py" (
    if exist "%cd%\generate_icon.py" (
        move /y "%cd%\generate_icon.py" "%cd%\generate_icon.py.backup" >nul 2>&1
    )
    move /y "%cd%\generate_icon_new.py" "%cd%\generate_icon.py" >nul 2>&1
    echo     ✓ 图标生成脚本已更新
)

echo ✓ 脚本更新检查完成
echo.

REM ===== 第四步：创建快捷方式 =====
echo [4/4] 正在创建桌面快捷方式...

REM 获取桌面路径
for /f "tokens=3" %%i in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Desktop') do (
    set "DESKTOP_PATH=%%i"
)

REM 获取脚本的完整路径
set "SCRIPT_PATH=%cd%\fm_searching.py"

REM 创建 VBScript 脚本来建立快捷方式
set "VBS_FILE=%temp%\create_shortcut.vbs"

(
    echo Set objWS = CreateObject("WScript.Shell"^)
    echo Set objLink = objWS.CreateShortcut("%DESKTOP_PATH%\防火门一键搜取.lnk"^)
    echo objLink.TargetPath = "python.exe"
    echo objLink.Arguments = "%SCRIPT_PATH%"
    echo objLink.WorkingDirectory = "%cd%"
    echo objLink.Description = "AutoCAD 防火门规格一键搜取工具"
    echo objLink.IconLocation = "%SCRIPT_PATH:fm_searching=fm_icon.ico%"
    echo objLink.Save
    echo MsgBox "快捷方式已创建到桌面！", 64, "创建成功"
) > %VBS_FILE%

cscript.exe %VBS_FILE% >nul 2>&1

if exist "%DESKTOP_PATH%\防火门一键搜取.lnk" (
    echo ✓ 桌面快捷方式创建成功！
) else (
    echo ✗ 快捷方式创建失败，但不影响正常使用
)

del /f /q %VBS_FILE% >nul 2>&1
echo.

REM ===== 完成 =====
echo ========================================
echo ✅ 环境配置完成！
echo ========================================
echo.
echo 📋 后续步骤：
echo   1. 在 AutoCAD 中打开你的 DWG 文件
echo   2. 双击桌面的"防火门一键搜取"快捷方式
echo   3. 脚本会自动扫描并生成 Excel 报表
echo.
echo 💡 提示：
echo   - 如有任何问题，请检查图层是否被冻结
echo   - 确保外部参照（XRef）文件已完整附加
echo   - 首次运行可能需要 AutoCAD 后台保持打开状态
echo.
pause
