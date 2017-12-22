echo off
xcopy /y E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.html E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.html
xcopy /y E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.css E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.css
xcopy /y E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.js E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.js
git add --all
git commit -m "any"
git push
set /p delExit=Press the ENTER key to exit...: