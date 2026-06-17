# WildPosting

PowerShell tooling for sorting wild-posting / street-poster photos. Reads
GPS-tagged photos from your phone, matches them to street folders with a
curated GPS-anchor system, and writes per-job sheets into an Excel workbook.

Runs on Windows + Microsoft Excel (uses Excel COM for the workbook side).

---

## What's in the box

| File | Purpose |
|---|---|
| `Create_poster_Pics.ps1` | Main orchestrator. Pops a dialog for City + Date, builds today's folder tree from the workbook, and sorts fresh photos by GPS proximity. |
| `LinkGenerator.ps1` | Standalone Street View hyperlink generator for the Master sheet. |
| `Compare-MasterList.ps1` | Reconciles the on-disk `<City>_Master\` folders against the workbook's Master sheet. Catches orphan rows / orphan folders / duplicates. |
| `Debug-PhotoGps.ps1` | Dumps EXIF GPS for any folder. Optional `-Sample <Street>` shows distance from each photo to that street's Master centroid - great for figuring out which Master pic is dragging the centroid the wrong way. |
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

1. Clone this repo somewhere convenient (the launcher assumes
   `%USERPROFILE%\WildPosting`, but you can edit `Launcher.vbs` if you put
   it elsewhere).
2. Make sure Microsoft Excel is installed (the scripts use Excel COM).
3. PowerShell 5.1 or newer (default on Windows 10/11 is fine).
4. Pick a city. Create the folder layout described in the working notes
   (`posters\<City>\<City>_Master\<Street>\`).
5. Create `<City>_Workbook.xlsx` with a `Master` sheet shaped as described.
6. Run `Compare-MasterList.ps1 -City <City>` to confirm folder names match
   workbook column C.
7. Run `Create_poster_Pics.ps1` and pick the city + date from the dialog.

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
