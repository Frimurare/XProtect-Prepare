# Förslag: Modularisera `Xprotect-prepare-19.ps1`

Det här förslaget delar upp huvudskriptet i tydliga moduler och ett litet startskript. Strukturen bevarar stöd för ps2exe genom att använda relativa sökvägar och hålla alla filer i samma katalogträd när exe-filen byggs. Den färdiga implementationen finns i `XProtect-Prepare-v2.ps1` och `modules/`.

## Föreslagen katalogstruktur
```
XProtect-Prepare/
├─ main.ps1                  # Liten meny/orkestrering, importerar moduler
├─ modules/
│  ├─ Core.psm1              # Banner, meny, gemensamma hjälpmetoder
│  ├─ Storage.psm1           # Diskkontroller, indexering, blockstorlek
│  ├─ Updates.psm1           # Windows Update-status, registerbackup
│  ├─ PowerAndServices.psm1  # Energischema, tjänstjusteringar
│  ├─ Network.psm1           # SNMP, NTP, brandväggsregler
│  ├─ Bloatware.psm1         # Avinstallation av bloatware, app-listor
│  └─ Hardening.psm1         # RDP, Defender-undantag, audit/logg-funktioner
└─ assets/                   # Ev. listfiler eller mallar
```

## Modulernas ansvar (exempel)
- **Core.psm1**
  - `Show-Banner`, `Show-MainMenu`, `Read-MenuChoice`, gemensam loggning och bekräftelsedialoger.
  - Initiering av loggmapp och kontroller som körs en gång (t.ex. administratörskontroll).
- **Storage.psm1**
  - `Get-AvailableDrives`, `Test-DriveBlockSize`, `Test-DriveIndexing`, `Disable-DriveIndexing` och körplan för disklayout.
- **Updates.psm1**
  - `Get-WindowsUpdateStatus`, `Set-WindowsUpdateStatus`, `New-RegistryBackup` och logik för att säkerställa korrekta update-policyer.
- **PowerAndServices.psm1**
  - `Set-PowerManagement` plus grupperade funktioner för att ändra tjänster och schemaläggare (t.ex. inspelningskritiska tjänster).
- **Network.psm1**
  - NTP-aktivering för kameror, SNMP-konfiguration, brandväggsregler för NTP och andra portar.
- **Bloatware.psm1**
  - Listor över oönskade appar och `Remove-…`-funktioner. Separera eventuellt OEM-specifika listor i underfunktioner.
- **Hardening.psm1**
  - RDP- och säkerhetsinställningar, Defender/AV-undantag, loggnings- och granskningsfunktioner.

## Entrypoint (`main.ps1`)
```powershell
# Kontrollera admin
. "$PSScriptRoot/modules/Core.psm1"
if (-not (Test-Administrator)) { Show-AdminWarning; exit }

# Ladda alla moduler
Get-ChildItem "$PSScriptRoot/modules" -Filter '*.psm1' | ForEach-Object {
    Import-Module $_.FullName -Force
}

Show-Banner
while ($true) {
    $choice = Show-MainMenu
    switch ($choice) {
        1 { Invoke-StorageMenu }
        2 { Invoke-UpdatesMenu }
        3 { Invoke-PowerAndServicesMenu }
        4 { Invoke-NetworkMenu }
        5 { Invoke-BloatwareMenu }
        6 { Invoke-HardeningMenu }
        'Q' { break }
    }
}
```
`Show-MainMenu` returnerar exempelvis ett tecken/nummer; varje modul exponerar en `Invoke-…Menu` som innehåller eventuella undermenyer (t.ex. separata val för blockstorlek eller indexering).

## Bygg med ps2exe
1. Placera `main.ps1` och `modules/` i samma mapp.
2. Kör `ps2exe .\main.ps1 -output Xprotect-prepare.exe` (alternativt `-icon` osv.).
3. Om du använder externa listfiler i `assets/`, ange `-include assets` eller bädda in listdata i modulerna.

## Fördelar
- **Underhållbarhet:** Varje modul har en begränsad uppsättning beroenden och kan testas separat.
- **Navigationsförbättring:** Undermenyer per modul minskar behovet av att starta om hela skriptet vid fel val.
- **Återanvändbarhet:** Flera miljöer kan dela moduler; `main.ps1` kan bytas ut utan att ändra funktionsblocken.
- **ps2exe-kompatibelt:** Relativa sökvägar och `Import-Module` gör att alla delar packas in i exe-filen utan att behöva ett monolitiskt skript.
