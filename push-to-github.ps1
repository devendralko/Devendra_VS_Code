$env:GIT_EDITOR = "cmd"
$env:GIT_MERGE_VERBOSITY = "1"

Set-Location "C:\Users\deven\OneDrive\Documentos\Learning\IDN tool"

# Fresh start
if (Test-Path .git) {
    Remove-Item -Recurse -Force .git
}

# Initialize
git init
git config user.email "deven@example.com"
git config user.name "Deven"
git config core.pager "cat"

# Add and commit
git add .
git commit -m "Initial commit: IDN/HCO Identification Pipeline Streamlit App"

# Push to GitHub
git remote add origin https://github.com/devendralko/Devendra_VS_Code.git
git branch -M main
git push -u origin main --force

Write-Host "Push completed!"
