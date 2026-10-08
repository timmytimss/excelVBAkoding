# Henter Excel-arbeidsboken Testfil.xlsm SPESIFIKT fra Running Object
# Table, uten a stole pa GetActiveObject("Excel.Application") (tvetydig
# nar flere Excel-prosesser kjorer samtidig - bekreftet 2026-09-23: en
# annen, EKTE Excel-prosess med Hakons Hovedfila-2027.xlsm kjorer
# parallelt og skal ALDRI roeres av denne modulen). Hele ROT-vandringen
# gjores i C# (ikke PowerShell) fordi PowerShell sin dynamiske COM-
# wrapping ikke handterer IRunningObjectTable/IEnumMoniker palitelig.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class RotFinder {
    public static object FindByNameSubstring(string substr) {
        IRunningObjectTable rot;
        GetRunningObjectTable(0, out rot);
        IEnumMoniker enumMoniker;
        rot.EnumRunning(out enumMoniker);
        enumMoniker.Reset();

        IMoniker[] monikers = new IMoniker[1];
        IntPtr fetched = IntPtr.Zero;
        IBindCtx bindCtx;
        CreateBindCtx(0, out bindCtx);

        while (enumMoniker.Next(1, monikers, fetched) == 0) {
            string displayName = null;
            try { monikers[0].GetDisplayName(bindCtx, null, out displayName); } catch { continue; }
            if (displayName != null && displayName.IndexOf(substr, StringComparison.OrdinalIgnoreCase) >= 0) {
                object obj;
                rot.GetObject(monikers[0], out obj);
                return obj;
            }
        }
        return null;
    }

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);
    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);
}
"@ -ErrorAction Stop

function Get-TestfilWorkbook {
    $wb = [RotFinder]::FindByNameSubstring("Testfil.xlsm")
    if ($null -eq $wb) { throw "Fant ikke Testfil.xlsm i Running Object Table." }
    if ($wb.Name -ne "Testfil.xlsm") { throw "SIKKERHETSSTOPP: fant '$($wb.Name)', ikke Testfil.xlsm. Avbryter." }
    return $wb
}

function Get-TestfilExcelApp {
    $wb = Get-TestfilWorkbook
    $excelApp = $wb.Application
    if ($excelApp.ActiveWorkbook.Name -ne "Testfil.xlsm" -and $wb.Name -ne "Testfil.xlsm") {
        throw "SIKKERHETSSTOPP: uventet arbeidsbok. Avbryter."
    }
    return $excelApp
}
