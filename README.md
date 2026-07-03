# SteamTools Install Manager Script

Windows PowerShell installer for **SkyTools**, **SteamTools**, and **OpenSteamTools**, plus **Millennium** and **Luatools** where applicable.

## Requirements

- PowerShell 5.1+ (built into Windows)

## Quick start

### Run from GitHub

```powershell
irm 'https://raw.githubusercontent.com/delabarra/SteamTools-Install-Manager-Script/main/run.ps1' | iex
```

### Bypass execution policy

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -NoExit -Command "irm 'https://raw.githubusercontent.com/delabarra/SteamTools-Install-Manager-Script/main/run.ps1' | iex"
```



## Temporary Windows Defender Exclusion (OpenSteamTools only)

Windows Defender flags the OpenSteamTools release zip, but not the contained DLLs. The script will ask permission to add a temporary exclusion so it can download and extract the files, then remove it when done. SkyTools and SteamTools do not need this.

## Download sources


| Component       | Source                                                                               |
| --------------- | ------------------------------------------------------------------------------------ |
| SkyTools DLLs   | [skyflare30.vercel.app/plugin/](https://skyflare30.vercel.app/plugin/)               |
| SteamTools DLLs | files.catbox.moe (pinned), update.steamcdn.com (fallback)                            |
| OpenSteamTools  | [OpenSteam001/OpenSteamTool](https://github.com/OpenSteam001/OpenSteamTool) releases |
| Millennium      | [SteamClientHomebrew/Millennium](https://github.com/SteamClientHomebrew/Millennium)  |
| Luatools        | [piqseu/ltsteamplugin](https://github.com/piqseu/ltsteamplugin)                      |
| STFixer         | [Selectively11/CloudRedirect](https://github.com/Selectively11/CloudRedirect) (`/stfixer`) |


