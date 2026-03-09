# C Drive Cleanup Tool

A compact Windows cleanup tool built with PowerShell + WPF.

## Features
- Sort cleanup items by impact level from high to low
- Show possible software or file-system impact before cleanup
- Provide one-click cleanup for selected items
- Estimate reclaimable disk space
- Skip locked or protected files and show results in a log panel

## Files
- `Clean-CDrive.ps1`: main UI and cleanup logic
- `Start-Cleanup-Tool.bat`: double-click launcher for Windows users

## Run
Double-click `Start-Cleanup-Tool.bat` or run:

```powershell
powershell -ExecutionPolicy Bypass -STA -File .\Clean-CDrive.ps1
```

For a non-UI smoke test:

```powershell
powershell -ExecutionPolicy Bypass -File .\Clean-CDrive.ps1 -NoUI
```

## Notes
- Running as Administrator is recommended for system cleanup paths.
- High-impact items are intentionally not selected by default.
- This project does not remove user documents directly; it focuses on temporary files and caches.

## Publish to GitHub
1. Install Git if it is not available on your machine.
2. Create a new GitHub repository.
3. In this folder, run:

```powershell
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```

If you want, this project can also be packaged into an `.exe` later with PS2EXE or a similar tool.
