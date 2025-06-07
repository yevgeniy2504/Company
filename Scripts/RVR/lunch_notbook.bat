@echo on
REM Включаем отображение всех команд (on вместо off)

REM Путь к виртуальному окружению (внутри кавычек — безопасно)
set "VENV_PATH=k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\PYTHON\.venv"
set "NOTEBOOK_PATH=k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\PYTHON\Scripts\task.ipynb"
set "PYTHON_PATH=%VENV_PATH%\Scripts\python.exe"

REM Проверим, существует ли интерпретатор
if not exist "%PYTHON_PATH%" (
    echo ERROR: Python interpreter not found at "%PYTHON_PATH%"
    pause
    exit /b
)

REM Активируем окружение (не обязательно, но можно)
call "%VENV_PATH%\Scripts\activate.bat"

REM Проверим, установлен ли Jupyter
"%PYTHON_PATH%" -m jupyter --version || (
    echo ERROR: Jupyter is not installed in this environment.
    pause
    exit /b
)

REM Запускаем Jupyter
"%PYTHON_PATH%" -m notebook "%NOTEBOOK_PATH%"

REM Оставим окно открытым
pause
