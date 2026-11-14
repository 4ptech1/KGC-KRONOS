@echo off
title KRONOS
pushd "%~dp0"
node -v >nul 2>nul
if not errorlevel 1 goto usenode
py -3 -c "pass" >nul 2>nul
if not errorlevel 1 goto usepy
python -c "pass" >nul 2>nul
if not errorlevel 1 goto usepython
goto usepowershell

:usenode
node "server\serve.mjs"
goto done

:usepy
py -3 "server\serve.py"
goto done

:usepython
python "server\serve.py"
goto done

:usepowershell
powershell -NoProfile -ExecutionPolicy Bypass -File "server\serve.ps1" "%CD%\app"
goto done

:done
popd
echo.
echo KRONOS has stopped. You can close this window.
pause
