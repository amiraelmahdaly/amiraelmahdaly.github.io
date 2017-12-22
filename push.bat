echo off
xcopy /s "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.html" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.html" /Y
xcopy /s "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.css" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.css" /Y
xcopy /s "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.js" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.js" /Y
git add --all
git commit -m "any"
git push
set /p delExit=Press the ENTER key to exit...: