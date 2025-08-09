@echo off

REM Cleaning and preparing deploy/src directory...
echo Cleaning and preparing deploy/src directory...
rd /s /q .\deploy\src
mkdir .\deploy\src

REM If deploy directory does not exist, create it and clone
if not exist .\deploy\Code.gs (
  if not exist .\deploy (
    echo Creating deploy directory...
    mkdir .\deploy
  )
  echo Reading PROJECTID from .env file...
  set PROJECTID=
  for /f "tokens=1,2 delims==" %%A in (.env) do (
    if "%%A"=="PROJECTID" set PROJECTID=%%B
  )
  REM Exit if PROJECTID is not set
  if "%PROJECTID%"=="" (
    echo PROJECTID not found in .env. Exiting deployment script.
    exit /b 1
  )
  echo PROJECTID found: %PROJECTID%
  echo Cloning Google Apps Script project with clasp in deploy directory...
  pushd .\deploy
  clasp clone %PROJECTID%
  REM Check if Code.gs contains runEngine, add if missing
  findstr /C:"function runEngine" Code.gs >nul
  if errorlevel 1 (
    echo Adding runEngine to Code.gs...
    echo.>>Code.gs
    echo function runEngine() {>>Code.gs
    echo   TrackerRun();>>Code.gs
    echo }>>Code.gs
  )
  popd
) else (
  echo Deploy directory already exists. Skipping clasp clone.
)

echo Copying src files to deploy/src...
xcopy src .\deploy\src /E

echo Changing to deploy directory...
pushd .\deploy

echo Pushing code to Google Apps Script project...
clasp push

popd

echo Deployment complete.