# SPED for Excel (v1.3.0)

An Excel add-in for accounting work: cell-format library, sheet tools, audit highlighter, data cleaning, compare tools, workbook inspector, sheet protection, favourites bar, run log with undo, and an update checker.

## Install
Run **SPED-Setup.exe**. It copies `SPED.xlam` to `%APPDATA%\Microsoft\AddIns`, switches it on in Excel and adds an entry to Windows *Apps > Installed apps* (uninstall from there, or run `SPED-Setup.exe /uninstall`).
Close Excel first. The installer is unsigned, so Windows SmartScreen may warn: choose *More info > Run anyway*.
Manual install: put `SPED.xlam` in the AddIns folder, then File > Options > Add-ins > Manage Excel Add-ins > Go > Browse.

## Updates
Settings > Check for updates reads a small text file: line 1 latest version, line 2 download address, line 3+ notes (see `sped_version.txt`).

## Layout
- `SPED.xlam`, `SPED-Setup.exe` - the built add-in and installer
- `src/` - VBA module sources (CRLF text, as exported from the project)
- `ribbon/` - ribbon XML and icons
- `installer/` - Go source of the installer (`GOOS=windows go build -ldflags "-s -w -H windowsgui"`; copy `SPED.xlam` next to `main.go` first)

Remembered sheet passwords are only lightly scrambled, not securely encrypted.
