@echo off
cd "C:\Users\deven\OneDrive\Documentos\Learning\IDN tool"

setx GIT_EDITOR cmd
setx GIT_PAGER cat

git init
git config user.email "deven@example.com"
git config user.name "Deven"

git add .
git commit -m "Initial commit: IDN/HCO Identification Pipeline"

git remote add origin https://github.com/devendralko/Devendra_VS_Code.git 2>NUL
git branch -M main
git push -u origin main --force

echo Push completed!
pause
