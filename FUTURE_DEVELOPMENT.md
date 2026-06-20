# Future Development — XProtect-Prepare

Backlog of planned improvements. Check off and move to the CHANGELOG when implemented.

---

## ✅ Done in v2.1 "Swedish Summer Edition"

- **Smart Client antivirus exclusions** — `Client.exe` + per-user `VideoOS` cache, conditional on the client being installed. (Without this, the antivirus scans thousands of fresh client files on first launch after a (re)install, making the client appear "broken" when it is merely being scanned in real time.)
- **VictoriaLogs antivirus exclusions** for Milestone 2026 R1 (with legacy Log Server kept for older versions).
- **MilestonePSTools** via the official web installer (2026 R1 compatible).
- **`-Silent` unattended mode**.

---

## 1. Refactor the antivirus exclusion logic into a single function (DRY)

The exclusion logic currently lives in **two places** with near-identical code: the standalone `Invoke-AntivirusStorageConfig` and the Complete Setup "Phase 10" block. Two copies tend to drift apart over time. Extract one function, e.g. `Set-MilestoneAvExclusions -Drives $x -IncludeSmartClient`, and call it from both menu paths. (v2.1 already shares the *new* additions via `Add-MilestoneClientAndLogExclusions`, which both paths call — the remaining duplication is the base folder/extension/process logic.)

## 2. Add a verification / report function for all three exclusion types

In the field, machines have been seen with only `ExclusionExtension` + `ExclusionProcess` set while `ExclusionPath` was empty — even though the tool had been "run on the drives" the day before.

- [ ] Add a verification function that reads and displays **all three**: `(Get-MpPreference).ExclusionPath / .ExclusionProcess / .ExclusionExtension`. Don't look at Path alone — the media database can be protected via *extensions* while Path is empty.
- [ ] Log clearly what was actually set (selected drives, folders, processes, extensions) so a later audit can confirm it.
- [ ] Also surface **policy / GPO exclusions** (`HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Exclusions`) — they do not always appear in `Get-MpPreference`.

## 3. Support for third-party antivirus

`Add-MpPreference` only configures Microsoft Defender. Consider detecting a third-party AV product (`root/SecurityCenter2` → `AntiVirusProduct`) and, at minimum, printing the exact folder/process/extension list the technician should add manually in that product.
