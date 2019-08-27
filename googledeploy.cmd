rd /s /q .\deployment\src
mkdir .\deployment\src
xcopy src .\deployment\src /E
cd ./deployment
clasp push
cd ../TradingEngine