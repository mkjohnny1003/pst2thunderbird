@echo off
echo ========================================================
echo   pst2thunderbird - GitHub Setup
echo   Target: https://github.com/mkjohnny1003/pst2thunderbird
echo ========================================================
echo.

git init
git add .
git commit -m "feat: initial commit of pst2thunderbird migration tool"
git branch -M main

echo.
echo Setting remote origin...
git remote remove origin 2>nul
git remote add origin https://github.com/mkjohnny1003/pst2thunderbird.git

echo.
echo ========================================================
echo Git init and commit complete!
echo Please create an empty repo named "pst2thunderbird"
echo on GitHub first, then run:
echo.
echo     git push -u origin main
echo ========================================================
pause
