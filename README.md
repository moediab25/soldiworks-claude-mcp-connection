# SolidWorks drawing automation

C# console tools that drive SolidWorks through its API. The main one, **MakeDrawings**, creates an
A4 detail drawing (`.SLDDRW` + `.pdf`) for every part in a project folder, with dimensions,
material and a filled-in title block.

## Requirements

- Windows with SolidWorks installed (tested with SolidWorks 2025). The projects reference the API
  DLLs at `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\`. If SolidWorks is installed
  somewhere else, change the `HintPath` in the `.csproj` files.
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

The drawing template and A4 sheet formats are found automatically in the newest
`C:\ProgramData\SOLIDWORKS\SOLIDWORKS <year>` folder that has them.

## Tools

| Folder | What it does |
|---|---|
| `SolidWorksConsole/MakeDrawings` | One A4 drawing per part, saved to `<project>\drawings\SLDDRW` and `<project>\drawings\PDF` |
| `SolidWorksConsole/InspectParts` | Lists each part in an assembly with its material, size and custom properties |
| `SolidWorksConsole` | Lists the components of the active assembly |

## MakeDrawings

```powershell
cd SolidWorksConsole\MakeDrawings
# every part listed in <project>\parts.txt (or every .SLDPRT in the folder if there is no parts.txt)
dotnet run -- --dir "D:\CAD\MyProject"
# just some parts
dotnet run -- --dir "D:\CAD\MyProject" Bracket Base_Plate
# save all open documents and close SolidWorks
dotnet run -- --save-and-exit
```

`--dir` defaults to the current folder. SolidWorks is started if it isn't running. While it works,
it keeps at most two documents of its own open (the part and its drawing). If a PDF is open in
another program, the new one is saved next to it as `<name>_new.pdf`.

### parts.txt

One part file name per line. `#` starts a comment. Leave out bought parts (motors, boards,
screens, fasteners) so they don't get drawings. See `SolidWorksConsole/examples/TER06_parts.txt`.

### What the drawings look like

- **Sheet:** A4, landscape or portrait, whichever fits the largest standard scale (5:1 … 1:20).
  Uses third-angle projection: front, top and right views plus an isometric.
- **Title block:** filled from the part's custom properties (title, description, drawn by, dates,
  finish, …). The material comes from the part.
- **Dimensions:** added edge-to-edge like Smart Dimension, so none read "0" and each one points at
  a real feature.
  - **Normal parts:** baseline dimensions from one outer edge, one number per row, with the
    overall size outermost.
  - **Plates with many holes:** each hole is measured from its nearest edge, and dimensions are
    stacked on the side of the view closest to their feature. Rows are packed so dimensions that
    don't overlap share a row.
  - **No repeats:** a size shown in one view isn't repeated in another.
- **Holes:**
  - One label per size: `4X Ø5.00`.
  - A tight regular grid gets one label: `9X Ø3.00 / 3 X 3 GRID, 8.00 PITCH`, and only its corner
    hole is located.
  - Labels sit in free space, with short leaders that don't cross.
- **Too-busy directions:** when a direction has too many positions to read on A4, the drawing shows
  overall sizes and adds the note "INTERNAL CUT-OUT POSITIONS PER CAD MODEL".

Set `DIMDEBUG=1` to print layout details while it runs.
