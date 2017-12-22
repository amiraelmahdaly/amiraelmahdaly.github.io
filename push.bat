echo off
xcopy "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.html" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.html" /h/e/k/f/c/y
xcopy "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.css" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.css" /h/e/k/f/c/y
xcopy "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\Home.js" "E:\Projects\Office Addins\ClaimMirroringAddin\ClaimMirroringAddinWeb\amiraelmahdaly.github.io\ClaimMirroringAddin\Home.js" /h/e/k/f/c/y
git add --all
git commit -m "any"
git push
set /p delExit=Press the ENTER key to exit...: