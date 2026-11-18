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
            swDoc = swApp.IActiveDoc2;
            if (swDoc == null)
            {
                Console.Error.WriteLine("No active document in SolidWorks.");
                return 3;
            }

            Console.WriteLine($"Active document: {swDoc.GetPathName() ?? swDoc.GetTitle()}");

            ExecuteTask(swDoc, swApp);

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
        var assy = swDoc as AssemblyDoc;
        if (assy == null) { Console.WriteLine("Active document is not an assembly: " + swDoc.GetTitle()); return; }
        Console.WriteLine("Assembly: " + swDoc.GetTitle());
        object[] comps = (object[])assy.GetComponents(false);
        if (comps == null) { Console.WriteLine("No components."); return; }
        var counts = new SortedDictionary<string, int>();
        foreach (object o in comps)
        {
            var c = (Component2)o;
            string path = c.GetPathName();
            string kind = path.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase) ? "ASM" : "PRT";
            Console.WriteLine($"{kind}\t{c.Name2}\t{(c.IsSuppressed() ? "suppressed" : "")}\t{path}");
            string key = kind + "\t" + Path.GetFileName(path);
            counts[key] = counts.TryGetValue(key, out int n) ? n + 1 : 1;
        }
        Console.WriteLine("\n--- Unique files (qty) ---");
        foreach (var kv in counts) Console.WriteLine($"{kv.Value}x\t{kv.Key}");
        Console.WriteLine($"Total instances: {comps.Length}");
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
