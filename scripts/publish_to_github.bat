@echo off
setlocal EnableExtensions
REM Erstellt/pusht das GitHub-Repo WAMAS-Lift-Store-Angebot (Windows CMD)
REM Voraussetzung: Git installiert. Optional: GitHub CLI (gh).

cd /d "%~dp0\.."
if not exist "server.py" (
  echo Bitte dieses Script aus dem Projektordner starten.
  echo Erwartet: server.py im Projektroot.
  exit /b 1
)

set "REPO_NAME=WAMAS-Lift-Store-Angebot"
set "OWNER=msobona"
set "VISIBILITY=private"
if not "%~1"=="" set "OWNER=%~1"

where git >nul 2>&1
if errorlevel 1 (
  echo Git fehlt. Bitte installieren: https://git-scm.com/download/win
  exit /b 1
)

where gh >nul 2>&1
if errorlevel 1 goto NO_GH

echo.
echo == GitHub CLI gefunden ==
gh auth status >nul 2>&1
if errorlevel 1 (
  echo Bitte einmal anmelden...
  gh auth login
)
echo Erstelle Repo %OWNER%/%REPO_NAME% ...
gh repo create "%OWNER%/%REPO_NAME%" --%VISIBILITY% --description "WAMAS Lift & Store Angebot – SSI SCHAEFER Kalkulator" --source=. --remote=origin --push
if errorlevel 1 (
  echo Repo existiert evtl. schon - pushe erneut...
  git remote remove origin 2>nul
  git remote add origin "https://github.com/%OWNER%/%REPO_NAME%.git"
  git push -u origin HEAD:master
  if errorlevel 1 git push -u origin HEAD:main
)
echo.
echo Fertig. Repo oeffnen:
gh repo view "%OWNER%/%REPO_NAME%" --web
exit /b 0

:NO_GH
echo.
echo GitHub CLI (gh) ist nicht installiert.
echo.
echo Mach bitte diese 2 Schritte:
echo.
echo 1) Im Browser leeres Repo anlegen:
echo    https://github.com/new
echo    Name: %REPO_NAME%
echo    Private, OHNE README / .gitignore / License
echo.
echo 2) Danach hier Enter druecken, dann wird gepusht.
echo.
pause

git remote remove origin 2>nul
git remote add origin "https://github.com/%OWNER%/%REPO_NAME%.git"
git branch -M master
git push -u origin master
if errorlevel 1 (
  echo.
  echo Push fehlgeschlagen. Pruefe:
  echo - Repo %OWNER%/%REPO_NAME% existiert und ist leer
  echo - Du bist bei Git angemeldet (GitHub Login / Credential Manager)
  echo - Alternativ: gh installieren https://cli.github.com/
  exit /b 1
)

echo.
echo Fertig: https://github.com/%OWNER%/%REPO_NAME%
start "" "https://github.com/%OWNER%/%REPO_NAME%"
exit /b 0
