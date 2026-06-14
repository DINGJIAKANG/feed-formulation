@echo off
echo ================================================
echo   饲料配方优化系统 - 启动脚本
echo ================================================
echo.

REM 检查是否存在虚拟环境
if exist .venv (
    echo ✓ 找到虚拟环境，正在激活...
    call .venv\Scripts\activate
) else (
    echo ⚠️ 未找到虚拟环境，将使用系统 Python
    echo 建议使用虚拟环境，运行以下命令创建：
    echo   python -m venv .venv
    echo   call .venv\Scripts\activate
    echo   pip install -r requirements.txt
    echo.
)

REM 启动 Streamlit 应用
echo ✓ 正在启动应用...
echo.
streamlit run streamlit_app.py --server.port 8501 --server.address localhost

pause
