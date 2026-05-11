using Microsoft.Win32;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using Word = Microsoft.Office.Interop.Word;

public class MacroCreator
{
    public static bool SetVbaTrustAccess(string appName, string version, bool enable = true)
    {
        string keyPath = $@"Software\Microsoft\Office\{version}\{appName}\Security";
        string valueName = "AccessVBOM";
        int valueData = enable ? 1 : 0;

        try
        {
            using (RegistryKey registryKey = Registry.CurrentUser.CreateSubKey(keyPath, true))
            {
                if (registryKey == null)
                {
                    Console.WriteLine("Failed to open or create registry key.");
                    return false;
                }

                registryKey.SetValue(valueName, valueData, RegistryValueKind.DWord);
            }

            string status = enable ? "Enabled" : "Disabled";
            Console.WriteLine($"Success: {status} VBA Trust Access for {appName} (Version {version})");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to modify registry: {ex.Message}");
            return false;
        }
    }

    private static void AddVbaCode(VBIDE.CodeModule codeModule, string vbaCode)
    {
        // Clear any existing content
        if (codeModule.CountOfLines > 0)
            codeModule.DeleteLines(1, codeModule.CountOfLines);

        string[] lines = vbaCode.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            codeModule.InsertLines(i + 1, lines[i]);
        }
    }

    public static string CreateMacros(string officeProduct, string vbaCode, string fileName, string vbaPassword = null)
    {
        officeProduct = officeProduct.ToLower();
        if (officeProduct != "excel" && officeProduct != "word")
            throw new ArgumentException("Supported products: Excel, Word");

        string filePath;

        if (officeProduct == "excel")
        {
            var app = new Excel.Application();
            app.Visible = false;
            var workbook = app.Workbooks.Add();

            VBIDE.VBComponent vbComponent = workbook.VBProject.VBComponents.Item("ThisWorkbook");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            // Apply VBA project protection
            if (!string.IsNullOrEmpty(vbaPassword))
            {
                ProtectVBAProject(workbook.VBProject, vbaPassword);
            }

            filePath = Path.GetFullPath(fileName + ".xls");
            workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
            workbook.Close(false);
            app.Quit();
        }
        else
        {
            var app = new Word.Application();
            app.Visible = false;
            var document = app.Documents.Add();

            VBIDE.VBComponent vbComponent = document.VBProject.VBComponents.Item("ThisDocument");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            // Apply VBA project protection
            if (!string.IsNullOrEmpty(vbaPassword))
            {
                ProtectVBAProject(document.VBProject, vbaPassword);
            }

            filePath = Path.GetFullPath(fileName + ".doc");
            document.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocument);
            document.Close(false);
            app.Quit();
        }

        return filePath;
    }

    public static string OpenMacros(string officeProduct, string filePath, string vbaCode, string vbaPassword = null)
    {
        officeProduct = officeProduct.ToLower();
        if (officeProduct != "excel" && officeProduct != "word")
            throw new ArgumentException("Supported products: Excel, Word");

        if (!File.Exists(filePath))
            throw new FileNotFoundException("File not found: " + filePath);

        if (officeProduct == "excel")
        {
            var app = new Excel.Application();
            app.Visible = true;
            var workbook = app.Workbooks.Open(filePath);

            VBIDE.VBComponent vbComponent = workbook.VBProject.VBComponents.Item("ThisWorkbook");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            // Apply VBA project protection
            if (!string.IsNullOrEmpty(vbaPassword))
            {
                ProtectVBAProject(workbook.VBProject, vbaPassword);
            }

            filePath = filePath + ".xls";
            workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
            workbook.Close(false);
            app.Quit();
        }
        else
        {
            var app = new Word.Application();
            app.Visible = true;
            var document = app.Documents.Open(filePath);

            VBIDE.VBComponent vbComponent = document.VBProject.VBComponents.Item("ThisDocument");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            // Apply VBA project protection
            if (!string.IsNullOrEmpty(vbaPassword))
            {
                ProtectVBAProject(document.VBProject, vbaPassword);
            }

            filePath = filePath + ".doc";
            document.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocument);
            document.Close(false);
            app.Quit();
        }

        return filePath;
    }

    private static void ProtectVBAProject(VBIDE.VBProject vbProject, string password)
    {
        var vbe = vbProject.VBE;

        vbe.ActiveVBProject = vbProject;
        vbProject.Name = "Project";

        vbe.MainWindow.Visible = true;
        System.Threading.Thread.Sleep(800);

        SetForegroundWindow(new IntPtr(vbe.MainWindow.HWnd));
        System.Threading.Thread.Sleep(800);

        // Try with TAB first
        bool success = TryProtect(vbe, vbProject, password, useTab: true);

        // If protection was not applied, retry without TAB
        if (!success)
        {
            SetForegroundWindow(new IntPtr(vbe.MainWindow.HWnd));
            System.Threading.Thread.Sleep(500);
            success = TryProtect(vbe, vbProject, password, useTab: false);
        }

        vbe.MainWindow.Visible = false;
    }

    private static bool TryProtect(VBIDE.VBE vbe, VBIDE.VBProject vbProject, string password, bool useTab)
    {
        // Open Tools → Project Properties
        System.Windows.Forms.SendKeys.SendWait("%T");
        System.Threading.Thread.Sleep(500);
        System.Windows.Forms.SendKeys.SendWait("E");
        System.Threading.Thread.Sleep(800);

        // Switch to Protection tab
        System.Windows.Forms.SendKeys.SendWait("^{TAB}");
        System.Threading.Thread.Sleep(500);

        // Some Office versions need TAB to focus the checkbox
        if (useTab)
        {
            System.Windows.Forms.SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(300);
        }

        // Space to tick "Lock project for viewing"
        System.Windows.Forms.SendKeys.SendWait(" ");
        System.Threading.Thread.Sleep(300);

        // Focus Password field via Alt+P
        string safePassword = EscapeSendKeys(password);
        System.Windows.Forms.SendKeys.SendWait("%P");
        System.Threading.Thread.Sleep(300);
        System.Windows.Forms.SendKeys.SendWait(safePassword);
        System.Threading.Thread.Sleep(300);

        // Focus Confirm password field via Alt+C
        System.Windows.Forms.SendKeys.SendWait("%C");
        System.Threading.Thread.Sleep(300);
        System.Windows.Forms.SendKeys.SendWait(safePassword);
        System.Threading.Thread.Sleep(300);

        // Click OK
        System.Windows.Forms.SendKeys.SendWait("{ENTER}");
        System.Threading.Thread.Sleep(800);

        // Check if protection was actually applied
        return vbProject.Protection == VBIDE.vbext_ProjectProtection.vbext_pp_locked;
    }

    private static string EscapeSendKeys(string input)
    {
        string[] specialChars = { "+", "^", "%", "~", "(", ")", "{", "}", "[", "]" };
        foreach (var ch in specialChars)
            input = input.Replace(ch, "{" + ch + "}");
        return input;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);


    static void Main()
    {
        string vbaCodeExcel = @"
Private Sub Workbook_Open()
    MsgBox ""Hello from Excel VBA""
End Sub
";
        string vbaCodeWord = @"
Private Sub Document_Open()
    MsgBox ""Hello from Word VBA""
End Sub
";      
        if (SetVbaTrustAccess("Excel", "16.0", true) && SetVbaTrustAccess("word", "16.0", true)) { CreateMacros("word", vbaCodeWord, "test_macro_word", vbaPassword: "MySecretPass123"); }
        
    }
}