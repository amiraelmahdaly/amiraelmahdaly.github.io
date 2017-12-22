echo off
git pull
xcopy  "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.html" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.html" /Y
xcopy  "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.css" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.css" /Y
xcopy  "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.js"  "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.js"  /Y
set /p delExit=Press the ENTER key to exit...: