# WildPosting

PowerShell + Python tooling for sorting wild-posting / street-poster photos.
Reads GPS-tagged photos from your phone, matches them to street folders with a
curated GPS-anchor system, and writes per-job sheets into an Excel workbook.
Includes a full UID / Street View / Lat-Lon maintenance pipeline.

Runs on Windows + Microsoft Excel (uses Excel COM for the PowerShell workbook side).

---

## Local machine paths (bipper9879)

These are the machine-specific paths this project uses by default.
Do **not** commit any file containing these paths — keep them in `.env` or
environment variables only.

| Purpose | Path |
|---|---|
| Posters root | `C:\Users\bippe\OneDrive\Documents\posters` |
| Philly workbook | `C:\Users\bippe\OneDrive\Documents\posters\Philly\Philly_Workbook.xlsx` |
| DC workbook | `C:\Users\bippe\OneDrive\Documents\posters\dc\DC_Workbook.xlsx` |
| Boston workbook | `C:\Users\bippe\OneDrive\Documents\posters\boston\Boston_Workbook.xlsx` |
| UID registry | `C:\Users\bippe\OneDrive\Documents\posters\Country_UID_Registry.xlsx` |
| UID log | `C:\Users\bippe\OneDrive\Documents\posters\UID_Automation_Log.txt` |
| Repo clone | `C:\Users\bippe\OneDrive\Workspace\Projects\WildPosting` |
| Launcher target | `%USERPROFILE%\WildPosting\Create_poster_Pics.ps1` |

To override the posters root without editing code, set an environment variable
before running any UID script:

```powershell
$env:WILDPOSTING_POSTERS_ROOT = "C:\your\custom\posters\root"
```

The UID scripts resolve root in this priority order:
1. `WILDPOSTING_POSTERS_ROOT` env var
2. `POSTERS_ROOT` env var
3. Current working directory (if it contains city subfolders or registry file)
4. `~\OneDrive\Documents\posters` (hardcoded fallback)

Check `UID_Automation_Log.txt` — the first line always prints `Active Root Directory`.
If that path is wrong, the sync will not find your city workbooks.

---

## What's in the box

| File | Purpose |
|---|---|
| `Create_poster_Pics.ps1` | Main orchestrator. Pops a dialog for City + Date, builds today's folder tree from the workbook, and sorts fresh photos by GPS proximity. |
| `Create_poster_Pics_variale_R.ps1` | Legacy launcher shim that forwards to `Create_poster_Pics.ps1` (safe to keep for old shortcuts). |
| `LinkGenerator.ps1` | Standalone Street View hyperlink generator for the Master sheet. |
| `Compare-MasterList.ps1` | Reconciles the on-disk `<City>_Master\` folders against the workbook's Master sheet. Catches orphan rows / orphan folders / duplicates. |
| `Debug-PhotoGps.ps1` | Dumps EXIF GPS for any folder. Optional `-Sample <Street>` shows distance from each photo to that street's Master centroid - great for figuring out which Master pic is dragging the centroid the wrong way. |
| `UID_Create.py` + `uid_utils.py` | Builds/refreshes `Country_UID_Registry.xlsx`, stamps `Site Code` UIDs back into each city `Master` tab, and writes `Lat`/`Lon` from Street View links. |
| `UID_Udate_sheets.py` | Updates latest dated city workbook(s) with UIDs + Lat/Lon values using `Country_UID_Registry.xlsx`. |
| `Run_UID_Maintenance.ps1` | Packed maintenance runner: compare folder/workbook names, refresh Master Street View links, stamp Master UID/Lat/Lon, then backfill dated files. |
| `PInMap.py` | Builds stacked pin map links from a selected dated sheet's Street View column. |
| `auto_fix_folders.py` | Auto-renames broken dated subfolder names against `<City>_Master`, exports unresolved cases to `needs_manual_fixing.xlsx`. |
| `sweep_by_folder.py` | Copies photos from dated folders back into matching `<City>_Master` subfolders and reports naming mismatches. |
| `Launcher.vbs` | Optional silent launcher for `Create_poster_Pics.ps1` (no PowerShell window). |
| `docs/Poster_Scripts_Working_Notes.md` | Detailed working notes covering algorithm, log format, urban-canyon GPS issues, and tuning. |

---

## How it works (short version)

1. You maintain a per-city workbook `<City>_Workbook.xlsx` with a `Master`
   sheet listing every poster location in column C.
2. You maintain a parallel `<City>_Master\<Street>\` folder with 2-6
   GPS-tagged anchor photos per location.
3. To run a job, you put a number in column D or E of the Master sheet for
   the rows you want to include, then run `Create_poster_Pics.ps1`.
4. The script:
   - Adds a dated sheet to the workbook with only the included rows.
   - Creates a folder tree for that date.
   - Loads anchors from each Master folder (averaged GPS = centroid).
   - For every photo in your inbox folder, computes distance to each
     centroid and moves it to the closest within a configurable radius.
   - Falls back to a timestamp-cluster check for photos with bad GPS.
   - Writes a detailed log so you can see what moved and why.

The full mechanics, knobs, and tuning advice live in
[docs/Poster_Scripts_Working_Notes.md](docs/Poster_Scripts_Working_Notes.md).

---

## Setup

1. Clone this repo:
   ```powershell
   git clone https://github.com/bipper9879/WildPosting.git
   cd WildPosting
   ```
2. Make sure Microsoft Excel is installed (the scripts use Excel COM).
3. PowerShell 5.1 or newer (default on Windows 10/11 is fine).  
   For the execution policy, run once per machine:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```
   Then unblock repo scripts after cloning:
   ```powershell
   Get-ChildItem *.ps1 | Unblock-File
   ```
4. If you use the Python helper tools, install dependencies once:
   ```powershell
   pip install pandas openpyxl pillow rapidfuzz tkcalendar
   ```
5. Create your posters folder layout (see working notes for details):
   ```
   C:\Users\<you>\OneDrive\Documents\posters\
     philly\
       Philly_Workbook.xlsx
       Philly_Master\
         <Street Name>\   ← 2–6 GPS-tagged anchor photos per street
     dc\
       DC_Workbook.xlsx
       DC_Master\
     boston\
       Boston_Workbook.xlsx
       Boston_Master\
     Country_UID_Registry.xlsx   ← auto-generated by UID_Create.py
   ```
6. Confirm folder names match workbook column C:
   ```powershell
   .\Compare-MasterList.ps1 -City Philly
   ```
7. Run the main sorter:
   ```powershell
   .\Create_poster_Pics.ps1
   ```

### One-command UID + Street View maintenance

Run this from the repo folder:

```powershell
.\Run_UID_Maintenance.ps1
```

Defaults:
- Processes `philly`, `dc`, `boston`
- Updates each city **Master** sheet every run
- Backfills **all** dated workbooks
- Skips dated workbook filenames containing `7-22` or `07-22`

Optional examples:

```powershell
# Only process latest dated workbook per city
.\Run_UID_Maintenance.ps1 -LatestOnly

# Custom cities and skip patterns
.\Run_UID_Maintenance.ps1 -Cities philly,dc -SkipPatterns "7-22","07-22","test"
```

If you hit `script not digitally signed` errors, run scripts via
`powershell.exe -ExecutionPolicy Bypass -File <script>.ps1` or set the
ExecutionPolicy to `RemoteSigned` for your user.

---

## Per-street radius overrides

Default match radius is 35 ft. Some streets in dense urban areas need more
because of GPS multipath drift. Drop a one-line text file named `.radius`
inside the Master subfolder, e.g.:

```
posters\DC\DC_Master\Connecticut Ave & Rhode Island Ave NW\.radius
```

with just an integer inside (e.g. `110`). The script logs every override
it loads as `[RADIUS] [Street] override = N ft`.

---

## License

MIT - see [LICENSE](LICENSE).
