@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

:: Double-click: keep window open even on errors
if /i not "%~1"=="run" (
    cmd /k call "%~f0" run
    exit /b 0
)

cd /d "%~dp0"

set REPO_NAME=interrogation_app
set BRANCH=main

echo ===================================================
echo   Interrogation App - Push to GitHub
echo   Repository: %REPO_NAME%
echo ===================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git is not installed or not in PATH.
    echo Download: https://git-scm.com/download/win
    goto :end
)

if not exist ".git" (
    echo [SETUP] Initializing local git repository...
    git init
    if errorlevel 1 goto :fail
    git branch -M %BRANCH%
    echo [OK] Local repository created.
    echo.
)

git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo [SETUP] No remote 'origin' configured yet.
    echo.
    echo Create an EMPTY repository on GitHub first:
    echo   https://github.com/new
    echo   Name: %REPO_NAME%
    echo   Do NOT add README / .gitignore / license
    echo.
    if exist ".github-remote.txt" (
        set /p GITHUB_REMOTE=<.github-remote.txt
        echo Using saved remote: !GITHUB_REMOTE!
    ) else (
        set /p GITHUB_USER=GitHub username: 
        if "!GITHUB_USER!"=="" (
            echo [ERROR] Username cannot be empty.
            goto :end
        )
        set "GITHUB_REMOTE=https://github.com/!GITHUB_USER!/!REPO_NAME!.git"
        echo !GITHUB_REMOTE!> .github-remote.txt
        echo Saved to .github-remote.txt
    )
    echo.
    git remote add origin "!GITHUB_REMOTE!"
    if errorlevel 1 (
        echo [ERROR] Failed to add remote.
        goto :end
    )
    echo [OK] Remote origin added.
    echo.
)

call :ensure_git_identity
if errorlevel 1 goto :end

set /p commit_msg=Commit message [Enter = Update from local]: 
if "%commit_msg%"=="" set "commit_msg=Update from local"

echo.
echo [1/3] Adding changes...
git add .
if errorlevel 1 goto :fail

echo.
echo [2/3] Committing...
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "%commit_msg%"
    if errorlevel 1 goto :fail
) else (
    echo No changes to commit. Trying push only...
)

echo.
echo [3/3] Pushing to origin %BRANCH%...
git rev-parse --abbrev-ref "@{u}" >nul 2>&1
if errorlevel 1 (
    git push -u origin %BRANCH%
) else (
    git push origin %BRANCH%
)

if errorlevel 0 (
    echo.
    echo ===================================================
    echo   SUCCESS! Pushed to GitHub.
    for /f "delims=" %%u in ('git remote get-url origin 2^>nul') do echo   %%u
    echo ===================================================
) else (
    echo.
    echo ===================================================
    echo   [ERROR] Push failed.
    echo   - Create empty repo %REPO_NAME% on GitHub
    echo   - Check .github-remote.txt
    echo   - Use Personal Access Token if login fails
    echo ===================================================
)

goto :end

:ensure_git_identity
set "GIT_USER_NAME="
set "GIT_USER_EMAIL="
for /f "delims=" %%n in ('git config user.name 2^>nul') do set "GIT_USER_NAME=%%n"
for /f "delims=" %%e in ('git config user.email 2^>nul') do set "GIT_USER_EMAIL=%%e"
if defined GIT_USER_NAME if defined GIT_USER_EMAIL exit /b 0

echo [SETUP] Git needs name and email for this repo only.
if exist ".git-author.txt" (
    for /f "usebackq tokens=1,2 delims=|" %%a in (".git-author.txt") do (
        set "GIT_USER_NAME=%%a"
        set "GIT_USER_EMAIL=%%b"
    )
    echo Using: !GIT_USER_NAME! / !GIT_USER_EMAIL!
) else (
    set /p GIT_USER_NAME=Your name: 
    set /p GIT_USER_EMAIL=Your email: 
    if "!GIT_USER_NAME!"=="" (
        echo [ERROR] Name required.
        exit /b 1
    )
    if "!GIT_USER_EMAIL!"=="" (
        echo [ERROR] Email required.
        exit /b 1
    )
    echo !GIT_USER_NAME!^|!GIT_USER_EMAIL!> .git-author.txt
)

git config user.name "!GIT_USER_NAME!"
git config user.email "!GIT_USER_EMAIL!"
if errorlevel 1 (
    echo [ERROR] Could not set git identity.
    exit /b 1
)
echo [OK] Git identity set.
exit /b 0

:fail
echo.
echo [ERROR] Operation failed.

:end
echo.
pause
endlocal
