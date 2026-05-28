$env:PATH += ";C:\Program Files\nodejs;$env:APPDATA\npm"
Set-Location $PSScriptRoot
clasp push --force
