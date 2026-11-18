using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksConsole;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        SldWorks? swApp = null;
        ModelDoc2? swDoc = null;

        try
        {
            swApp = ConnectToSolidWorks();
            if (swApp == null)
            {
                Console.Error.WriteLine("Unable to attach to SolidWorks.");
                return 2;
            }

            swApp.Visible = true;
            ExecuteTask(null!, swApp); // task opens the assembly by path

            Console.WriteLine("Task completed successfully.");
            return 0;
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"COM error: 0x{ex.ErrorCode:X8} - {ex.Message}");
            return ex.ErrorCode;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
        finally
        {
            if (swDoc != null) Marshal.ReleaseComObject(swDoc);
            if (swApp != null) Marshal.ReleaseComObject(swApp);
        }
    }

    private static void ExecuteTask(ModelDoc2 swDoc, SldWorks swApp)
    {
        // assembly file to inspect: first command-line argument
        string AsmPath = Path.GetFullPath(System.Environment.GetCommandLineArgs().Skip(1).FirstOrDefault() ?? throw new Exception("usage: dotnet run -- <assembly.SLDASM>"));
        int e0 = 0, w0 = 0;
        swDoc = (ModelDoc2)swApp.OpenDoc6(AsmPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref e0, ref w0);
        var assy = swDoc as AssemblyDoc;
        if (assy == null) { Console.WriteLine("Active document is not an assembly: " + swDoc.GetTitle()); return; }
        Console.WriteLine("Assembly: " + swDoc.GetTitle());
        object[] comps = (object[])assy.GetComponents(true);
        var paths = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (object o in comps) paths.Add(((Component2)o).GetPathName());
        foreach (string p in paths)
        {
            if (p.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase)) { Console.WriteLine($"{Path.GetFileName(p)}\tASSEMBLY"); continue; }
            int err = 0, warn = 0;
            var doc = swApp.OpenDoc6(p, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn) as ModelDoc2;
            if (doc == null) { Console.WriteLine($"OPENFAIL\t{Path.GetFileName(p)}\terr={err}"); continue; }
            string db;
            string mat = ((PartDoc)doc).GetMaterialPropertyName2("", out db);
            var cpm = doc.Extension.CustomPropertyManager[""];
            var names = (object[]?)cpm.GetNames();
            var props = new List<string>();
            if (names != null)
                foreach (string n in names)
                {
                    cpm.Get6(n, false, out string v, out string r, out _, out _);
                    props.Add($"{n}={r}");
                }
            var box = (double[])((PartDoc)doc).GetPartBox(true);
            string size = box == null ? "?" : $"{(box[3]-box[0])*1000:F1}x{(box[4]-box[1])*1000:F1}x{(box[5]-box[2])*1000:F1}mm";
            Console.WriteLine($"{Path.GetFileName(p)}\tmat=[{mat}]\t{size}\t{string.Join("; ", props)}");
        }
    }

    private static SldWorks? ConnectToSolidWorks()
    {
        try
        {
            return Interaction.GetObject(string.Empty, "SldWorks.Application") as SldWorks;
        }
        catch (COMException)
        {
        }

        var type = Type.GetTypeFromProgID("SldWorks.Application");
        return type == null ? null : (SldWorks?)Activator.CreateInstance(type);
    }
}

