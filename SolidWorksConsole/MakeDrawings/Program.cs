using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace MakeDrawings;

// Creates one A4 drawing (SLDDRW + PDF) per manufactured top-level part.
// Usage: dotnet run -- [--dir <project folder>] [partFileName ...]
//        dotnet run -- --save-and-exit
internal static class Program
{
    // Set in Main: project folder (--dir, default = current folder) and SolidWorks template paths
    // (newest installed SolidWorks version is detected, so this runs on any PC with SolidWorks).
    static string Dir = "", Template = "", SheetFormat = "", SheetFormatPortrait = "";
    static readonly HashSet<string> Created = new(); // titles of drawings this run made (for cleanup after a failure)
    const double SheetW = 0.297, SheetH = 0.210;

    // Parts to draw: names on the command line, else parts.txt in the project folder
    // (one part file name per line, # = comment), else every .SLDPRT in the folder.
    static string[] LoadParts(string[] names)
    {
        if (names.Length > 0) return names;
        string list = Path.Combine(Dir, "parts.txt");
        if (File.Exists(list))
            return File.ReadAllLines(list).Select(l => l.Split('#')[0].Trim()).Where(l => l != "")
                       .Select(l => Path.GetFileNameWithoutExtension(l)).ToArray();
        return Directory.GetFiles(Dir, "*.SLDPRT").Select(Path.GetFileNameWithoutExtension).ToArray()!;
    }

    // Drawing template + A4 sheet formats from the newest C:\ProgramData\SOLIDWORKS\SOLIDWORKS <year> folder.
    static void FindTemplates(SldWorks swApp)
    {
        string root = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData), "SOLIDWORKS");
        string data = (Directory.Exists(root) ? Directory.GetDirectories(root, "SOLIDWORKS *") : Array.Empty<string>())
            .Where(d => File.Exists(Path.Combine(d, "templates", "Drawing.DRWDOT")) && File.Exists(Path.Combine(d, "lang", "english", "sheetformat", "a4 - portrait.slddrt")))
            .OrderByDescending(d => d).FirstOrDefault()
            ?? throw new Exception($"no SolidWorks drawing template found under {root}");
        Template = Path.Combine(data, "templates", "Drawing.DRWDOT");
        SheetFormat = Path.Combine(data, "lang", "english", "sheetformat", "a4 - landscape.slddrt");
        SheetFormatPortrait = Path.Combine(data, "lang", "english", "sheetformat", "a4 - portrait.slddrt");
        foreach (var f in new[] { SheetFormat, SheetFormatPortrait })
            if (!File.Exists(f)) throw new Exception($"missing SolidWorks file: {f}");
    }

    static readonly (int num, int den)[] Scales =
        { (5, 1), (2, 1), (1, 1), (1, 2), (1, 5), (1, 10), (1, 20) };

    [STAThread]
    private static int Main(string[] args)
    {
        SldWorks? swApp = null;
        if (args.Length > 0 && args[0] == "--save-and-exit") return SaveAndExit();
        int di = Array.IndexOf(args, "--dir");
        Dir = Path.GetFullPath(di >= 0 && di + 1 < args.Length ? args[di + 1] : Directory.GetCurrentDirectory());
        if (di >= 0) args = args.Where((_, k) => k != di && k != di + 1).ToArray();
        if (!Directory.Exists(Dir)) { Console.Error.WriteLine($"Folder not found: {Dir}"); return 2; }
        try
        {
            swApp = ConnectToSolidWorks();
            if (swApp == null) { Console.Error.WriteLine("Unable to attach to SolidWorks."); return 2; }
            swApp.Visible = true;
            FindTemplates(swApp);

            string outDir = Path.Combine(Dir, "drawings");
            Directory.CreateDirectory(Path.Combine(outDir, "SLDDRW"));
            Directory.CreateDirectory(Path.Combine(outDir, "PDF"));
            CloseLeftovers(swApp);
            var todo = LoadParts(args);
            if (todo.Length == 0) { Console.Error.WriteLine($"No parts to draw in {Dir}"); return 2; }
            int ok = 0;
            foreach (string p in todo)
            {
                // Each part opens itself + one drawing; close both afterwards so SolidWorks never
                // holds more than those 2 (plus whatever the user already had open).
                var before = OpenDocs(swApp).Select(d => d.GetTitle()).ToHashSet();
                string partPath = Path.Combine(Dir, p + ".SLDPRT");
                Created.Clear();
                try { MakeDrawing(swApp, p, outDir); ok++; }
                catch (Exception ex) { Console.WriteLine($"FAIL\t{p}\t{ex.Message}"); }
                finally
                {
                    // only what this run opened: the part itself and drawings it created (never a file the user opened meanwhile)
                    try
                    {
                        foreach (var d in OpenDocs(swApp).Where(d => !before.Contains(d.GetTitle())
                                     && (Created.Contains(d.GetTitle()) || string.Equals(d.GetPathName(), partPath, StringComparison.OrdinalIgnoreCase))))
                            swApp.CloseDoc(d.GetTitle());
                    }
                    catch (COMException) { } // SolidWorks went away; next part reconnects
                }
            }
            Console.WriteLine($"Done: {ok}/{todo.Length} drawings in {outDir}");
            return ok == todo.Length ? 0 : 1;
        }
        catch (Exception ex) { Console.Error.WriteLine($"Error: {ex.Message}"); return 1; }
        finally { if (swApp != null) Marshal.ReleaseComObject(swApp); }
    }

    static void MakeDrawing(SldWorks swApp, string name, string outDir)
    {
        string partPath = Path.Combine(Dir, name + ".SLDPRT");
        int err = 0, warn = 0;
        var part = swApp.OpenDoc6(partPath, (int)swDocumentTypes_e.swDocPART,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn) as ModelDoc2
            ?? throw new Exception($"open part failed err={err}");

        // Part properties -> copied to the drawing so $PRP and $PRPSHEET title-block fields both resolve.
        var props = new List<(string k, string v)>();
        var pcpm = part.Extension.CustomPropertyManager[""];
        foreach (string n in (object[]?)pcpm.GetNames() ?? Array.Empty<object>())
        {
            pcpm.Get6(n, false, out _, out string r, out _, out _);
            props.Add((n, r));
        }
        string mat = ((PartDoc)part).GetMaterialPropertyName2("", out _);
        if (string.IsNullOrEmpty(mat)) throw new Exception("no material assigned");

        var box = (double[])((PartDoc)part).GetPartBox(true);
        double dx = box[3] - box[0], dy = box[4] - box[1], dz = box[5] - box[2];

        // Third angle: front bottom-left, top above it, right beside it, isometric in the free corner.
        // 1) Probe drawing: place the views once just to count how many dimension rows each needs.
        const double m = 0.004; // SolidWorks view outline margin
        bool plate = false; // many-hole plate: hole positions split to the nearest side of the view
        DimCounts Probe(double ps)
        {
            var probe = (ModelDoc2)swApp.NewDocument(Template, (int)swDwgPaperSizes_e.swDwgPaperA4size, SheetW, SheetH);
            Created.Add(probe.GetTitle());
            var pv = new[]
            {
                Place((DrawingDoc)probe, partPath, "*Front", ps, 0.07, 0.10),
                Place((DrawingDoc)probe, partPath, "*Top", ps, 0.07, 0.17),
                Place((DrawingDoc)probe, partPath, "*Right", ps, 0.20, 0.10),
            };
            probe.ForceRebuild3(false); // view outlines are stale until the sheet rebuilds
            var c = SmartDim(swApp, probe, pv, null, SheetW, SheetH, false, dryRun: true, holeTable: plate);
            swApp.CloseDoc(probe.GetTitle());
            return c;
        }

        // 2) Space between/around views so every row fits, then pick the orientation + largest ISO
        //    scale for which views + rows fit above/beside the title block.
        var zones = new[]
        {
            (portrait: false, xL: 0.014, yB: 0.072, w: 0.269, h: 0.124), // landscape, above title block
            (portrait: false, xL: 0.014, yB: 0.014, w: 0.089, h: 0.182), // landscape, left of title block
            (portrait: true,  xL: 0.014, yB: 0.072, w: 0.182, h: 0.211), // portrait, above title block
        };
        double holeBand = 0, left = 0, gapFR = 0, gapFT = 0, aboveTop = 0, spareW = 0, spareH = 0;
        var L = zones[2]; (int num, int den) sc = (1, 50);
        void Pick(DimCounts c)
        {
            // hole labels: one row under the front view per 2 hole sizes
            holeBand = 0.008 + (c.HoleSizes[0] + 1) / 2 * 0.0075 + (c.HBelow > 0 ? c.HBelow * RowH + Pad : 0);
            left = Math.Max(c.VCols[0], c.VCols[1]) * RowH + Pad + 0.004;
            gapFR = (c.VCols[2] + c.VRight) * RowH + Pad + 0.010 + (c.VRight > 0 ? Pad : 0);
            // between front and top: the front view's horizontal rows plus the top view's hole-label rows
            gapFT = (c.HRows[0] + (c.HoleSizes[1] + 1) / 2) * RowH + Pad + 0.012;
            aboveTop = c.HRows[1] * RowH + Pad;
            L = zones[2]; sc = (1, 50);
            var zs = zones;
            foreach (var scale in Scales)
            {
                double s0 = (double)scale.num / scale.den;
                double needW = left + dx * s0 + 2 * m + gapFR + dz * s0 + 2 * m;
                double needH = holeBand + dy * s0 + 2 * m + gapFT + dz * s0 + 2 * m + aboveTop;
                var fit = zs.Where(z => needW <= z.w && needH <= z.h).ToList();
                if (fit.Count > 0) { L = fit[0]; sc = scale; spareW = L.w - needW; spareH = L.h - needH; break; }
            }
        }
        // Row counts depend on scale (lines closer than MinGap on paper merge), so re-count at a larger
        // chosen scale; counts from a larger probe scale are safe for any smaller final scale.
        double probed = 0.05;
        var cnt = Probe(probed);
        if (cnt.Busy && cnt.FrontHoles >= 4 && cnt.FrontHoles <= 40) { plate = true; cnt = Probe(probed); }
        while (true)
        {
            Pick(cnt);
            double chosen = (double)sc.num / sc.den;
            if (chosen <= probed + 1e-9) break;
            cnt = Probe(chosen); probed = chosen;
        }
        var (num, den) = sc;
        double s = (double)num / den;

        double sw = L.portrait ? SheetH : SheetW, sh = L.portrait ? SheetW : SheetH;
        int paper = (int)(L.portrait ? swDwgPaperSizes_e.swDwgPaperA4sizeVertical : swDwgPaperSizes_e.swDwgPaperA4size);
        var drw = (ModelDoc2)swApp.NewDocument(Template, paper, sw, sh);
        Created.Add(drw.GetTitle());
        var dd = (DrawingDoc)drw;
        string sheetName = ((Sheet)dd.GetCurrentSheet()).GetName();
        dd.SetupSheet5(sheetName, paper, (int)swDwgTemplates_e.swDwgTemplateCustom,
            num, den, false, L.portrait ? SheetFormatPortrait : SheetFormat, sw, sh, "Default", true);

        double x0 = L.xL + left + spareW / 2, y0 = L.yB + holeBand + spareH / 2;
        double frontR = x0 + dx * s + 2 * m, frontT = y0 + dy * s + 2 * m;
        double rightL = frontR + gapFR, rightR = rightL + dz * s + 2 * m;
        var ortho = new[]
        {
            Place(dd, partPath, "*Front", s, x0 + m + dx * s / 2, y0 + m + dy * s / 2),
            Place(dd, partPath, "*Top", s, x0 + m + dx * s / 2, frontT + gapFT + m + dz * s / 2),
            Place(dd, partPath, "*Right", s, rightL + m + dz * s / 2, y0 + m + dy * s / 2),
        };

        // Isometric gets its own scale in the larger free zone: above the right view's rows, or right of it.
        double xmax = sw - 0.014, ymax = sh - 0.014;
        double isoW0 = 0.87 * (dx + dz), isoH0 = 0.5 * (dx + dz) + 0.82 * dy;
        var isoZones = new[] { (x: rightL, y: frontT + cnt.HRows[2] * RowH + Pad + 0.006), (x: rightR + 0.012, y: 0.078) };
        var best = isoZones.Select(z => (z, fit: Math.Min((xmax - z.x - 2 * m) / isoW0, (ymax - z.y - 2 * m) / isoH0)))
                           .OrderByDescending(t => t.fit).First();
        var iso = PickScale(best.fit);
        var isoView = Place(dd, partPath, "*Isometric", (double)iso.num / iso.den, (best.z.x + xmax) / 2, (best.z.y + ymax) / 2);

        // First view's auto-scaling resets the sheet scale; put it back so the title block SCALE is right.
        ((Sheet)dd.GetCurrentSheet()).SetScale(num, den, false, false);

        drw.ClearSelection2(true);
        drw.ForceRebuild3(false);
        bool busy = SmartDim(swApp, drw, ortho, isoView, sw, sh, L.portrait, dryRun: false, holeTable: plate).Busy;
        if (busy)
        {
            // Too many steps/cut-outs to dimension legibly: sheet shows overall sizes + hole sizes.
            var note = (Note)drw.InsertNote("NOTE: INTERNAL CUT-OUT POSITIONS PER CAD MODEL");
            ((Annotation)note.GetAnnotation()).SetPosition2(L.xL + 0.01, L.yB - 0.0, 0);
        }

        // Title-block text: the template's title box only fits ~3 short lines.
        foreach (Note n in (object[]?)((SolidWorks.Interop.sldworks.View)dd.GetFirstView()).GetNotes() ?? Array.Empty<object>())
        {
            string link = n.PropertyLinkedText ?? "";
            string desc = props.FirstOrDefault(p => p.k == "Description").v ?? "";
            double h = link.Contains("\"Description\"") ? (desc.Length > 30 ? 0.003 : 0.0035)
                     : link.Contains("File Name") ? 0.004 : 0;
            if (h == 0) continue;
            var a = (Annotation)n.GetAnnotation();
            var tf = (TextFormat)a.GetTextFormat(0);
            tf.CharHeight = h;
            a.SetTextFormat(0, false, tf);
        }

        // Weight/Mass are already pulled from the part ($PRPSHEET); copying them doubles the text.
        var dcpm = drw.Extension.CustomPropertyManager[""];
        foreach (var (k, v) in props.Where(p => p.k != "Weight" && p.k != "Mass"))
            dcpm.Add3(k, (int)swCustomInfoType_e.swCustomInfoText, v, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
        if (!props.Any(p => p.k == "Material"))
            dcpm.Add3("Material", (int)swCustomInfoType_e.swCustomInfoText, mat, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);

        drw.ClearSelection2(true); // selected dims print highlighted in the PDF
        drw.ForceRebuild3(false);
        string drwPath = Path.Combine(outDir, "SLDDRW", name + ".SLDDRW");
        string pdfPath = Path.Combine(outDir, "PDF", name + ".pdf");
        if (IsLocked(pdfPath)) { pdfPath = Path.Combine(outDir, "PDF", name + "_new.pdf"); Console.WriteLine($"  {name}.pdf is open in another program, saved as {name}_new.pdf"); }
        if (!drw.Extension.SaveAs3(drwPath, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn))
            throw new Exception($"save SLDDRW failed err={err}");
        if (!drw.Extension.SaveAs3(pdfPath, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn))
            throw new Exception($"save PDF failed err={err}");
        swApp.CloseDoc(drw.GetTitle());
        Console.WriteLine($"OK\t{name}\t{(L.portrait ? "portrait" : "landscape")}\tscale {num}:{den}\tiso {iso.num}:{iso.den}{(plate ? "\tplate" : "")}\tmat {mat}");
    }

    static bool IsLocked(string f)
    {
        if (!File.Exists(f)) return false;
        try { using var _ = File.Open(f, FileMode.Open, FileAccess.ReadWrite, FileShare.None); return false; }
        catch (IOException) { return true; }
    }

    // Smart-Dimension style: every dimension is added between two chosen edges (same as clicking
    // them with Smart Dimension), so there are no "0" labels and each number points at a feature.
    // Per view and direction, every step edge and hole centre is measured from the same outer edge
    // (baseline), one row per number with the overall size outermost, so dimension lines are always
    // horizontal/vertical and numbers never share a row. A direction showing exactly the same
    // positions as another view on the same model axis is skipped. One diameter per hole size,
    // "nX Ø.." when repeated; hole labels go to the first free spot around the view.
    // A direction with more than MaxFeatures positions shows only its overall size.
    // dryRun = only count rows (used to space the views before dimensioning).
    const int MaxFeatures = 7;
    const int MaxLanes = 10; // plate mode: most packed dim rows per side
    const double MinGap = 0.0015; // smallest distance on paper between two dimensioned lines
    const double TextH = 0.0055, CharW = 0.0027, RowH = 0.0075, Pad = 0.007; // text box incl. leader shoulder
    record Feat(Edge E, double Pos, bool Line, double Perp = 0);
    record struct Box(double X0, double Y0, double X1, double Y1)
    {
        public static Box At(double cx, double cy, double w, double h) => new(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2);
        public bool Hits(Box b) => X0 < b.X1 && b.X0 < X1 && Y0 < b.Y1 && b.Y0 < Y1;
    }
    record DimCounts(int[] HRows, int[] VCols, int[] HoleSizes, bool Busy, int FrontHoles, int HBelow, int VRight);

    static DimCounts SmartDim(SldWorks swApp, ModelDoc2 drw, SolidWorks.Interop.sldworks.View[] ortho,
                              SolidWorks.Interop.sldworks.View? iso, double sw, double sh, bool portrait, bool dryRun, bool holeTable = false)
    {
        var mu = (MathUtility)swApp.GetMathUtility();
        string[] hAxis = { "X", "X", "Z" }, vAxis = { "Y", "Z", "Y" };
        var shownSets = new HashSet<string>();
        int[] hRows = new int[3], vCols = new int[3], holes = new int[3];
        bool busy = false;
        int added = 0, nudged = 0, frontHoles = 0, hBelow = 0, vRight = 0;

        var taken = new List<Box>();
        if (!dryRun)
        {
            foreach (var vw in iso == null ? ortho : ortho.Append(iso))
            {
                var ol = (double[])vw.GetOutline();
                taken.Add(new Box(ol[0] + 0.003, ol[1] + 0.003, ol[2] - 0.003, ol[3] - 0.003));
            }
            taken.Add(new Box(portrait ? 0 : 0.105, 0, sw, 0.070));
        }
        var leaders = new List<(double x0, double y0, double x1, double y1)>();
        static bool Crosses((double x0, double y0, double x1, double y1) a, (double x0, double y0, double x1, double y1) b)
        {
            static double Side(double ax, double ay, double bx, double by, double px, double py) => (bx - ax) * (py - ay) - (by - ay) * (px - ax);
            return Side(a.x0, a.y0, a.x1, a.y1, b.x0, b.y0) * Side(a.x0, a.y0, a.x1, a.y1, b.x1, b.y1) < 0
                && Side(b.x0, b.y0, b.x1, b.y1, a.x0, a.y0) * Side(b.x0, b.y0, b.x1, b.y1, a.x1, a.y1) < 0;
        }
        bool Free(Box b) => b.X0 > 0.014 && b.Y0 > 0.014 && b.X1 < sw - 0.014 && b.Y1 < sh - 0.014 && !taken.Any(t => t.Hits(b));

        for (int i = 0; i < ortho.Length; i++)
        {
            var v = ortho[i];
            var xf = v.ModelToViewTransform;
            double sc = v.ScaleDecimal;
            var o = (double[])v.GetOutline();
            double[] S(double[] p) => (double[])((MathPoint)((MathPoint)mu.CreatePoint(new[] { p[0], p[1], p[2] })).MultiplyTransform(xf)).ArrayData;

            var vert = new List<Feat>(); var horiz = new List<Feat>();
            var circles = new List<(Edge e, double x, double y, double dMm)>();
            foreach (var e in ((object[]?)v.GetVisibleEntities2(null, (int)swViewEntityType_e.swViewEntityType_Edge) ?? Array.Empty<object>()).OfType<Edge>())
            {
                var c = (Curve)e.GetCurve();
                var sv = e.GetStartVertex() as Vertex; var ev = e.GetEndVertex() as Vertex;
                if (c.IsLine() && sv != null && ev != null)
                {
                    var a = S((double[])sv.GetPoint()); var b = S((double[])ev.GetPoint());
                    double dx = b[0] - a[0], dy = b[1] - a[1];
                    if (Math.Abs(dx) < 1e-6 && Math.Abs(dy) > 1e-5) vert.Add(new Feat(e, (a[0] + b[0]) / 2, true, (a[1] + b[1]) / 2));
                    else if (Math.Abs(dy) < 1e-6 && Math.Abs(dx) > 1e-5) horiz.Add(new Feat(e, (a[1] + b[1]) / 2, true, (a[0] + b[0]) / 2));
                }
                else if (c.IsCircle() && (sv == null || sv == ev))
                {
                    var cp = (double[])c.CircleParams;
                    // only circles seen face-on in this view (edge-on ones look like lines)
                    var ax = (double[])((MathVector)((MathVector)mu.CreateVector(new[] { cp[3], cp[4], cp[5] })).MultiplyTransform(xf)).ArrayData;
                    double norm = 0.99 * Math.Sqrt(ax[0] * ax[0] + ax[1] * ax[1] + ax[2] * ax[2]);
                    var ctr = S(cp);
                    if (Math.Abs(ax[2]) >= norm) circles.Add((e, ctr[0], ctr[1], Math.Round(cp[6] * 2000, 2)));
                    // seen edge-on (e.g. a tube end) the circle draws as a straight line: use it as a step edge
                    else if (Math.Abs(ax[1]) >= norm) horiz.Add(new Feat(e, ctr[1], true, ctr[0]));
                    else if (Math.Abs(ax[0]) >= norm) vert.Add(new Feat(e, ctr[0], true, ctr[1]));
                }
            }

            // Hole table: tags every hole (A1, A2, B1...) and lists X/Y from the part's bottom-left corner plus
            // size, so a many-hole plate is fully located without a forest of leaders.
            // Selection marks: 1 = datum vertex, 2 = hole edges.
            List<(double Key, List<(Edge e, double x, double y, double dMm)> Items, string Below)> Grids()
            {
                var res = new List<(double, List<(Edge e, double x, double y, double dMm)>, string)>();
                foreach (var g in circles.GroupBy(c => c.dMm))
                {
                    // single-link clusters: holes closer than 4 mm on paper
                    var left = g.ToList();
                    while (left.Count > 0)
                    {
                        var cl = new List<(Edge e, double x, double y, double dMm)> { left[0] }; left.RemoveAt(0);
                        for (int k = 0; k < cl.Count; k++)
                        {
                            var near = left.Where(c => Math.Abs(c.x - cl[k].x) < 0.004 && Math.Abs(c.y - cl[k].y) < 0.004).ToList();
                            cl.AddRange(near); left.RemoveAll(near.Contains);
                        }
                        // distinct model-mm positions, values within 0.5 mm counted as one (float noise)
                        List<double> Bins(IEnumerable<double> v)
                        {
                            var r = new List<double>();
                            foreach (var q in v.Select(p => p / sc * 1000).OrderBy(p => p)) if (r.Count == 0 || q - r[^1] > 0.5) r.Add(q);
                            return r;
                        }
                        var xs = Bins(cl.Select(c => c.x));
                        var ys = Bins(cl.Select(c => c.y));
                        bool Even(List<double> q) => q.Zip(q.Skip(1), (a, b) => b - a).All(d => Math.Abs(d - (q[1] - q[0])) < 0.2);
                        if (cl.Count < 4 || xs.Count < 2 || ys.Count < 2 || xs.Count * ys.Count != cl.Count || !Even(xs) || !Even(ys)) continue;
                        double px = Math.Round(xs[1] - xs[0], 2), py = Math.Round(ys[1] - ys[0], 2);
                        var items = cl.OrderBy(c => c.x + c.y).ToList(); // first = bottom-left corner hole
                        res.Add((g.Key, items, $"{xs.Count} X {ys.Count} GRID, {(Math.Abs(px - py) < 0.05 ? $"{px:0.00}" : $"{px:0.00} X {py:0.00}")} PITCH"));
                    }
                }
                return res;
            }

            var dimHoles = circles; // holes whose positions get dimensioned
            // the band a stack of n dim rows occupies beside the view (dimension + extension lines)
            void Reserve(bool horizontal, bool farSide, int n)
            {
                if (n == 0) return;
                double d = Pad + (n - 1) * RowH + TextH / 2, e = 0.004; // e: extension-line overshoot past the view
                taken.Add(horizontal
                    ? (farSide ? new Box(o[0] - e, o[1] - d, o[2] + e, o[1]) : new Box(o[0] - e, o[3], o[2] + e, o[3] + d))
                    : (farSide ? new Box(o[2], o[1] - e, o[2] + d, o[3] + e) : new Box(o[0] - d, o[1] - e, o[0], o[3] + e)));
            }

            DisplayDimension? Dim(Edge a, Edge b, double x, double y)
            {
                drw.ClearSelection2(true);
                v.SelectEntity(a, false);
                if (b != a) v.SelectEntity(b, true);
                var d = drw.AddDimension2(x, y, 0) as DisplayDimension;
                if (d != null) { ((Annotation)d.GetAnnotation()).SetPosition2(x, y, 0); added++; }
                return d;
            }

            // Baseline dims from the view's left/bottom edge, one row per number, overall size outermost.
            // split (many-hole plate) -> PlateDirection.
            int Direction(List<Feat> lines, Func<(Edge e, double x, double y, double dMm), double> holePos,
                          Func<(Edge e, double x, double y, double dMm), double> holePerp, string axis, bool horizontal, bool split, out int farCount)
            {
                farCount = 0;
                if (lines.Count < 2) return 0;
                if (split) return PlateDirection(lines, holePos, holePerp, axis, horizontal, out farCount);
                var baseF = lines.OrderBy(f => f.Pos).First();
                double Mm(Feat f) => Math.Round((f.Pos - baseF.Pos) / sc * 1000, 1);
                var feats = lines.Concat(dimHoles.Select(c => new Feat(c.e, holePos(c), false)))
                                 .Where(f => Mm(f) >= 0.5).GroupBy(Mm).Select(g => g.OrderByDescending(f => f.Line).First())
                                 .OrderBy(Mm).ToList();
                // Lines closer than MinGap on paper can't be told apart (sheet-thickness offsets, bend lines):
                // keep the overall size and drop lines crowding a kept line or the base edge. Holes are always kept.
                var kept = new List<Feat>();
                for (int k = feats.Count - 1; k >= 0; k--)
                    if (kept.Count == 0 || !feats[k].Line || (kept[^1].Pos - feats[k].Pos >= MinGap && feats[k].Pos - baseF.Pos >= MinGap))
                        kept.Add(feats[k]);
                kept.Reverse(); feats = kept;
                // a size already shown on this model axis in another view is not repeated
                feats = feats.Where(f => !shownSets.Contains(axis + Mm(f))).ToList();
                if (feats.Count == 0) return 0;
                if (feats.Count > MaxFeatures) { busy = true; feats = new() { feats.Last() }; }
                foreach (var f in feats) shownSets.Add(axis + Mm(f));
                if (dryRun) return feats.Count;
                for (int k = 0; k < feats.Count; k++)
                {
                    var f = feats[k];
                    double mid = (baseF.Pos + f.Pos) / 2, w = Mm(f).ToString("0.00").Length * CharW;
                    double x = horizontal ? mid : o[0] - Pad - k * RowH;
                    double y = horizontal ? o[3] + Pad + k * RowH : mid;
                    var b = horizontal ? Box.At(x, y, w, TextH) : Box.At(x, y, TextH, w);
                    if (!Free(b)) Console.WriteLine($"  WARNING: {Mm(f):0.00} may touch another item");
                    taken.Add(b);
                    Dim(baseF.E, f.E, x, y);
                }
                Reserve(horizontal, false, feats.Count);
                return feats.Count;
            }

            // Many-hole plate, the way a drafter stacks it:
            //  - every position is measured from the NEAREST outer edge (left/right, bottom/top), so values are
            //    short and dims from opposite edges can sit in the same row/column;
            //  - a dim goes on the side of the view nearest to its feature (above/below, left/right);
            //  - rows/columns are packed: a dim takes the innermost row whose dims (and numbers) it doesn't overlap;
            //  - the overall size is outermost on the default side (above / left).
            int PlateDirection(List<Feat> lines, Func<(Edge e, double x, double y, double dMm), double> holePos,
                               Func<(Edge e, double x, double y, double dMm), double> holePerp, string axis, bool horizontal, out int farCount)
            {
                var lo = lines.OrderBy(f => f.Pos).First(); var hi = lines.OrderBy(f => f.Pos).Last();
                double total = Math.Round((hi.Pos - lo.Pos) / sc * 1000, 1);
                double Abs(Feat f) => Math.Round((f.Pos - lo.Pos) / sc * 1000, 1);
                double mid0 = horizontal ? (o[1] + o[3]) / 2 : (o[0] + o[2]) / 2;
                double perpLo = horizontal ? o[1] : o[0], perpHi = horizontal ? o[3] : o[2];
                bool FarP(double perp) => horizontal ? perp < mid0 - 1e-6 : perp > mid0 + 1e-6;

                var cands = lines.Where(f => f != lo && f != hi)
                                 .Concat(dimHoles.Select(c => new Feat(c.e, holePos(c), false, holePerp(c))))
                                 .Where(f => Abs(f) >= 0.5 && Abs(f) <= total - 0.5).ToList();
                var spans = new List<(Feat r, Feat f, double val, bool far, double a, double b)>();
                foreach (var g in cands.GroupBy(Abs))
                {
                    // one dim per position; it points at the member nearest the perpendicular edge it's drawn from
                    var rep = g.OrderByDescending(f => f.Line)
                               .ThenBy(f => Math.Round(Math.Min(Math.Abs(f.Perp - perpLo), Math.Abs(f.Perp - perpHi)) * 1e4)) // 0.1 mm bins
                               .ThenBy(f => f.Perp).First(); // ties: lower/left member, so probe and final agree
                    if (!shownSets.Add(axis + g.Key)) continue; // already shown in another view
                    bool fromLo = g.Key <= total / 2;
                    var r = fromLo ? lo : hi;
                    if (rep.Line && Math.Abs(rep.Pos - r.Pos) < MinGap) continue; // line too close to its edge to read
                    double val = fromLo ? g.Key : Math.Round(total - g.Key, 1);
                    spans.Add((r, rep, val, FarP(rep.Perp), Math.Min(r.Pos, rep.Pos), Math.Max(r.Pos, rep.Pos)));
                }
                shownSets.Add(axis + total);

                // pack into lanes per side; the number's length counts as part of the span
                var lanes = new[] { new List<List<(double a, double b)>>(), new List<List<(double a, double b)>>() };
                var placed = new List<(Feat r, Feat f, double val, bool far, int lane)>();
                foreach (var s in spans.OrderBy(s => s.b - s.a))
                {
                    double half = s.val.ToString("0.00").Length * CharW / 2 + 0.0012, m = (s.a + s.b) / 2;
                    double a = Math.Min(s.a, m - half), b = Math.Max(s.b, m + half);
                    var l = lanes[s.far ? 1 : 0];
                    int lane = l.FindIndex(q => q.All(z => z.b <= a || b <= z.a));
                    if (lane < 0) { l.Add(new()); lane = l.Count - 1; }
                    l[lane].Add((a, b));
                    placed.Add((s.r, s.f, s.val, s.far, lane));
                }
                if (lanes[0].Count > MaxLanes || lanes[1].Count > MaxLanes)
                {
                    busy = true; placed.Clear(); lanes[0].Clear(); lanes[1].Clear();
                }
                farCount = lanes[1].Count;
                int nearCount = lanes[0].Count + 1; // + overall
                if (System.Environment.GetEnvironmentVariable("DIMDEBUG") == "1") Console.WriteLine($"  plate {(horizontal ? "H" : "V")} dry={dryRun} near={nearCount} far={lanes[1].Count} spans={spans.Count}");
                if (dryRun) return nearCount;

                void Label(Feat r, Feat f, double mm, bool farSide, int k)
                {
                    double mid = (r.Pos + f.Pos) / 2;
                    double x = horizontal ? mid : farSide ? o[2] + Pad + k * RowH : o[0] - Pad - k * RowH;
                    double y = horizontal ? (farSide ? o[1] - Pad - k * RowH : o[3] + Pad + k * RowH) : mid;
                    double w = mm.ToString("0.00").Length * CharW;
                    var bx = horizontal ? Box.At(x, y, w, TextH) : Box.At(x, y, TextH, w);
                    if (!Free(bx)) Console.WriteLine($"  WARNING: {mm:0.00} may touch another item");
                    taken.Add(bx);
                    Dim(r.E, f.E, x, y);
                }
                foreach (var p in placed) Label(p.r, p.f, p.val, p.far, p.lane);
                Label(lo, hi, total, false, lanes[0].Count);
                Reserve(horizontal, false, lanes[0].Count + 1); Reserve(horizontal, true, lanes[1].Count);
                return nearCount;
            }

            var groups = circles.GroupBy(c => c.dMm).Select(g => (Key: g.Key, Items: g.ToList(), Below: "")).ToList();
            if (i == 0) frontHoles = circles.Count;
            bool split = holeTable && i == 0;
            if (split)
            {
                // Tight hole grids: one label "nX Ø, a X b GRID, p PITCH" and only the grid's corner hole
                // is located; the other holes of each size get the usual "nX Ø" label.
                var grids = Grids();
                var inGrid = grids.SelectMany(g => g.Items).ToHashSet();
                var skip = grids.SelectMany(g => g.Items.Skip(1)).ToHashSet();
                groups = circles.Where(c => !inGrid.Contains(c)).GroupBy(c => c.dMm)
                                .Select(g => (Key: g.Key, Items: g.ToList(), Below: "")).Concat(grids).ToList();
                dimHoles = circles.Where(c => !skip.Contains(c)).ToList();
            }
            hRows[i] = Direction(vert, c => c.x, c => c.y, hAxis[i], true, split, out int far0);
            vCols[i] = Direction(horiz, c => c.y, c => c.x, vAxis[i], false, split, out int far1);
            if (split) { hBelow = far0; vRight = far1; }
            holes[i] = groups.Count;
            if (dryRun) continue;
            foreach (var g in groups)
            {
                int n = g.Items.Count;
                double th = g.Below == "" ? TextH : 2 * TextH;
                double w = Math.Max(((n > 1 ? $"{n}X " : "") + "Ø" + g.Key.ToString("0.00")).Length, g.Below.Length) * CharW;
                // Spots in rows under the view and in a column right of it; the label takes the free spot
                // closest to one of its holes and points at that hole, so leaders stay short and don't cross the part.
                var cands = new List<(double x, double y)>();
                for (int r = 0; r < 5; r++)
                    for (double cx = o[0] - 0.06; cx <= o[2] + 0.06; cx += 0.004)
                        cands.Add((cx, o[1] - 0.006 - r * 0.0075));
                for (double cy = o[1] - 0.02; cy < o[3]; cy += 0.0075) cands.Add((o[2] + 0.005 + w / 2, cy));
                // anywhere else on the sheet as a last resort (still nearest-to-hole first)
                for (double cx = 0.02; cx < sw - 0.02; cx += 0.004)
                    for (double cy = 0.02; cy < sh - 0.02; cy += 0.004) cands.Add((cx, cy));
                var free = cands.Where(p => Free(Box.At(p.x, p.y, w, th))).ToList();
                if (free.Count == 0) Console.WriteLine("  WARNING: hole label may touch another item");
                var pairs = (free.Count > 0 ? free : cands.Take(1).ToList())
                    .SelectMany(p => g.Items.Select(c => (p, c)))
                    .OrderBy(t => Math.Pow(t.p.x - t.c.x, 2) + Math.Pow(t.p.y - t.c.y, 2)).ToList();
                // prefer a leader that doesn't cross one already drawn
                var (spot, h) = pairs.FirstOrDefault(t => !leaders.Any(l => Crosses(l, (t.c.x, t.c.y, t.p.x, t.p.y))), pairs[0]);
                leaders.Add((h.x, h.y, spot.x, spot.y));
                taken.Add(Box.At(spot.x, spot.y, w, th));
                var d = Dim(h.e, h.e, spot.x, spot.y);
                // horizontal text on a broken leader, so the label occupies exactly the box reserved for it
                d?.SetBrokenLeader2(false, (int)swDisplayDimensionLeaderText_e.swBrokenLeaderHorizontalText);
                if (d != null && n > 1)
                    d.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, $"{n}X " + d.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix));
                if (d != null && g.Below != "") d.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutBelow, g.Below);
            }
        }
        drw.ClearSelection2(true);
        if (!dryRun) Console.WriteLine($"  smart dims added: {added}{(busy ? " (busy directions: overall size only)" : "")}");
        return new DimCounts(hRows, vCols, holes, busy, frontHoles, hBelow, vRight);
    }

    // Saves every open document that has unsaved changes, then shuts SolidWorks down.
    // Attaches only to a running session; never launches one.
    static int SaveAndExit()
    {
        SldWorks swApp;
        try { swApp = (SldWorks)Interaction.GetObject(string.Empty, "SldWorks.Application"); }
        catch (Exception) { Console.WriteLine("SolidWorks is not running."); return 0; }
        int failed = 0;
        foreach (var d in OpenDocs(swApp).Where(d => d.GetSaveFlag()))
        {
            int e = 0, w = 0;
            bool ok = d.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref e, ref w);
            Console.WriteLine($"{(ok ? "saved" : "SAVE FAILED")}\t{d.GetTitle()}{(ok ? "" : $" err={e}")}");
            if (!ok) failed++;
        }
        if (failed > 0) { Console.WriteLine("Not closing SolidWorks: some files did not save."); return 1; }
        swApp.ExitApp();
        Console.WriteLine("SolidWorks closed.");
        return 0;
    }

    static List<ModelDoc2> OpenDocs(SldWorks swApp) =>
        ((object[]?)swApp.GetDocuments() ?? Array.Empty<object>()).OfType<ModelDoc2>().ToList();

    // Close what earlier runs left open: unmodified parts/drawings from the project folder and unsaved
    // empty drawings. Assemblies and anything with unsaved changes are left alone.
    static void CloseLeftovers(SldWorks swApp)
    {
        int closed = 0;
        foreach (var d in OpenDocs(swApp))
        {
            string path = d.GetPathName() ?? "";
            bool mine = path.StartsWith(Dir, StringComparison.OrdinalIgnoreCase)
                        && (path.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".SLDDRW", StringComparison.OrdinalIgnoreCase));
            bool emptyDrawing = path == "" && d.GetType() == (int)swDocumentTypes_e.swDocDRAWING;
            if ((mine || emptyDrawing) && !d.GetSaveFlag()) { swApp.CloseDoc(d.GetTitle()); closed++; } // CloseDoc discards changes
        }
        Console.WriteLine($"Closed {closed} leftover document(s); {OpenDocs(swApp).Count} still open.");
    }

    static void Stage(string s)
    {
        if (System.Environment.GetEnvironmentVariable("DIMDEBUG") == "1") Console.WriteLine("   > " + s);
    }

    static (int num, int den) PickScale(double fit)
    {
        var sc = Scales.FirstOrDefault(t => (double)t.num / t.den <= fit);
        return sc.num == 0 ? (1, 50) : sc;
    }

    static SolidWorks.Interop.sldworks.View Place(DrawingDoc dd, string model, string viewName, double scale, double x, double y)
    {
        var v = dd.CreateDrawViewFromModelView3(model, viewName, x, y, 0) as SolidWorks.Interop.sldworks.View
            ?? throw new Exception($"view {viewName} failed");
        v.ScaleDecimal = scale;
        v.Position = new[] { x, y };
        return v;
    }

    private static SldWorks? ConnectToSolidWorks()
    {
        try { return Interaction.GetObject(string.Empty, "SldWorks.Application") as SldWorks; }
        catch (COMException) { }
        var type = Type.GetTypeFromProgID("SldWorks.Application");
        return type == null ? null : (SldWorks?)Activator.CreateInstance(type);
    }
}
