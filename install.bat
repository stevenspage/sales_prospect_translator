@echo off
chcp 65001 >nul
echo ========================================
echo     Sales Prospector Translator
echo         依赖库一键安装脚本
echo ========================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误：未找到Python环境！
    echo 请先安装Python 3.8或更高版本
    echo 下载地址：https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo ✅ Python环境检查通过
echo.

echo 开始安装依赖库...
echo ========================================

echo 📦 安装核心依赖...
pip install pandas==2.3.2
if errorlevel 1 echo ❌ pandas安装失败

pip install openpyxl==3.1.5
if errorlevel 1 echo ❌ openpyxl安装失败

pip install requests==2.32.4
if errorlevel 1 echo ❌ requests安装失败

pip install beautifulsoup4==4.13.5
if errorlevel 1 echo ❌ beautifulsoup4安装失败

pip install chardet==5.2.0
if errorlevel 1 echo ❌ chardet安装失败

echo.
echo 📄 安装文档生成依赖...
pip install newspaper3k==0.2.8
if errorlevel 1 echo ❌ newspaper3k安装失败

pip install fpdf==1.7.2
if errorlevel 1 echo ❌ fpdf安装失败

pip install python-docx==1.1.2
if errorlevel 1 echo ❌ python-docx安装失败

echo.
echo 🔧 安装可选增强功能依赖...
pip install trafilatura==2.0.0
if errorlevel 1 echo ❌ trafilatura安装失败

pip install langdetect==1.0.9
if errorlevel 1 echo ❌ langdetect安装失败

pip install google-search-results==2.4.2
if errorlevel 1 echo ❌ google-search-results安装失败

echo.
echo ========================================
echo 安装完成！
echo
echo.
echo 🚀 现在可以运行 sales_prospector_translator.py 了，本页面可以手动关闭
echo ========================================
echo.
pause
