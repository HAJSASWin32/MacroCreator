using Microsoft.Win32;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using Word = Microsoft.Office.Interop.Word;

// Requires NuGet package: OpenMcdf (https://www.nuget.org/packages/OpenMcdf)
using OpenMcdf;
public class MacroCreator
{
    public static bool SetVbaTrustAccess(string appName, string version, bool enable = true)
    {
        string keyPath = $@"Software\Microsoft\Office\{version}\{appName}\Security";
        string valueName = "AccessVBOM";
        int valueData = enable ? 1 : 0;

        try
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath, true))
            {
                if (key == null)
                {
                    Console.WriteLine("Failed to open or create registry key.");
                    return false;
                }
                key.SetValue(valueName, valueData, RegistryValueKind.DWord);
            }

            Console.WriteLine($"Success: {(enable ? "Enabled" : "Disabled")} VBA Trust Access for {appName} (Version {version})");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to modify registry: {ex.Message}");
            return false;
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
            var app = new Excel.Application { Visible = false };
            var workbook = app.Workbooks.Add();

            VBIDE.VBComponent vbComponent =
                workbook.VBProject.VBComponents.Item("ThisWorkbook");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            filePath = Path.GetFullPath(fileName + ".xls");
            workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
            workbook.Close(false);
            app.Quit();
        }
        else
        {
            var app = new Word.Application { Visible = false };
            var document = app.Documents.Add();

            VBIDE.VBComponent vbComponent = document.VBProject.VBComponents.Item("ThisDocument");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            filePath = Path.GetFullPath(fileName + ".doc");
            document.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocument);
            document.Close(false);
            app.Quit();
        }

        
        if (!string.IsNullOrEmpty(vbaPassword))
            ProtectVbaProjectInFile(filePath, vbaPassword);

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
            var app = new Excel.Application { Visible = true };
            var workbook = app.Workbooks.Open(filePath);

            VBIDE.VBComponent vbComponent =
                workbook.VBProject.VBComponents.Item("ThisWorkbook");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            filePath = filePath + ".xls";
            workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
            workbook.Close(false);
            app.Quit();
        }
        else
        {
            var app = new Word.Application { Visible = true };
            var document = app.Documents.Open(filePath);

            VBIDE.VBComponent vbComponent =
                document.VBProject.VBComponents.Item("ThisDocument");
            AddVbaCode(vbComponent.CodeModule, vbaCode);

            filePath = filePath + ".doc";
            document.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocument);
            document.Close(false);
            app.Quit();
        }

        if (!string.IsNullOrEmpty(vbaPassword))
            ProtectVbaProjectInFile(filePath, vbaPassword);

        return filePath;
    }

    private static void AddVbaCode(VBIDE.CodeModule codeModule, string vbaCode)
    {
        if (codeModule.CountOfLines > 0)
            codeModule.DeleteLines(1, codeModule.CountOfLines);

        string[] lines = vbaCode.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
        for (int i = 0; i < lines.Length; i++)
            codeModule.InsertLines(i + 1, lines[i]);
    }

    public static void ProtectVbaProjectInFile(string filePath, string password)
    {
        string ext = Path.GetExtension(filePath).ToLower();

        if (ext == ".xlsm" || ext == ".docm" || ext == ".pptm")
            ProtectVbaProjectInOoxml(filePath, password);
        else
            ProtectVbaProjectInCfb(filePath, password); 
    }

    private static void ProtectVbaProjectInOoxml(string filePath, string password)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        try
        {
            ZipFile.ExtractToDirectory(filePath, tempDir);

            string vbaBin = Directory.GetFiles(tempDir, "vbaProject.bin",
                                               SearchOption.AllDirectories)
                                     .FirstOrDefault()
                ?? throw new FileNotFoundException(
                       "vbaProject.bin not found inside the Office package.");

            ProtectVbaProjectInCfb(vbaBin, password);

            File.Delete(filePath);
            ZipFile.CreateFromDirectory(tempDir, filePath);
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    private static void ProtectVbaProjectInCfb(string cfbPath, string password)
    {
        using (var cf = new CompoundFile(cfbPath, CFSUpdateMode.Update, CFSConfiguration.Default))
        {
            CFStream projectStream = FindStream(cf.RootStorage, "PROJECT")
                ?? throw new InvalidOperationException(
                       "PROJECT stream not found in CFB file. Ensure the file was saved with a VBA project.");

            byte[] rawBytes = projectStream.GetData();
            Encoding enc = Encoding.GetEncoding(1252);
            string projectTxt = enc.GetString(rawBytes);

            string projectId = ExtractProjectId(projectTxt);
            projectTxt = RewriteProtectionFields(projectTxt, password, projectId);

            projectStream.SetData(enc.GetBytes(projectTxt));
            cf.Commit();
        }
    }

    private static string RewriteProtectionFields(string text, string password, string originalProjectId)
    {
        const string ZeroId = "{00000000-0000-0000-0000-000000000000}";
        byte projKey = ComputeProjKey(ZeroId);

        string cmg = EncryptProtectionState(projKey, userProtected: true);
        string dpb = EncryptPassword(projKey, password);
        string gc = EncryptVisibilityState(projKey, visible: true);

        var lines = text
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
            .ToList();

        for (int i = lines.Count - 1; i >= 0; i--)
        {
            string trimmed = lines[i].Trim();
            if (trimmed.StartsWith("ID=", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("CMG=", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("DPB=", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("GC=", StringComparison.OrdinalIgnoreCase))
            {
                lines.RemoveAt(i);
            }
        }

        int insertAt = lines.FindIndex(l =>
            l.TrimStart().StartsWith("[Host Extender Info]",
                StringComparison.OrdinalIgnoreCase) ||
            l.TrimStart().StartsWith("[Workspace]",
                StringComparison.OrdinalIgnoreCase));

        if (insertAt < 0) insertAt = lines.Count;

        lines.Insert(insertAt, $"GC=\"{gc}\"");
        lines.Insert(insertAt, $"DPB=\"{dpb}\"");
        lines.Insert(insertAt, $"CMG=\"{cmg}\"");
        lines.Insert(insertAt, $"ID=\"{ZeroId}\"");

        return string.Join("\r\n", lines);
    }

    private static string DataEncrypt(byte projKey, byte[] data, byte seed = 0)
    {

        var output = new System.Collections.Generic.List<byte>();

        byte versionEnc = (byte)(seed ^ 0x02); 
        byte projKeyEnc = (byte)(seed ^ projKey);

        byte unencryptedByte1 = projKey;
        byte encryptedByte1 = projKeyEnc;
        byte encryptedByte2 = versionEnc;

        output.Add(seed);
        output.Add(versionEnc);
        output.Add(projKeyEnc);

        int ignoredLength = (seed & 6) / 2;
        for (int i = 0; i < ignoredLength; i++)
        {
            byte tempValue = 0; 
            byte byteEnc = (byte)(tempValue ^ (encryptedByte2 + unencryptedByte1));
            output.Add(byteEnc);
            encryptedByte2 = encryptedByte1;
            encryptedByte1 = byteEnc;
            unencryptedByte1 = tempValue;
        }

        uint length = (uint)data.Length;
        byte[] lenBytes =
        {
            (byte)(length         & 0xFF),
            (byte)((length >>  8) & 0xFF),
            (byte)((length >> 16) & 0xFF),
            (byte)((length >> 24) & 0xFF),
        };
        foreach (byte b in lenBytes)
        {
            byte byteEnc = (byte)(b ^ (encryptedByte2 + unencryptedByte1));
            output.Add(byteEnc);
            encryptedByte2 = encryptedByte1;
            encryptedByte1 = byteEnc;
            unencryptedByte1 = b;
        }

        foreach (byte dataByte in data)
        {
            byte byteEnc = (byte)(dataByte ^ (encryptedByte2 + unencryptedByte1));
            output.Add(byteEnc);
            encryptedByte2 = encryptedByte1;
            encryptedByte1 = byteEnc;
            unencryptedByte1 = dataByte;
        }

        return BytesToDoubledHex(output.ToArray());
    }

    private static byte[] BuildHashDataStructure(string password)
    {
        byte[] key = { 0x38, 0x7A, 0xB1, 0x5C };

        byte[] pwBytes = Encoding.GetEncoding(1252).GetBytes(password);
        byte[] toHash = pwBytes.Concat(key).ToArray();
        byte[] pwHash;
        using (var sha1 = SHA1.Create())
            pwHash = sha1.ComputeHash(toHash);   

        EncodeNulls(key, out uint grbitKey, out byte[] keyNoNulls);
        EncodeNulls(pwHash, out uint grbitHashNull, out byte[] hashNoNulls);

        byte gByte0 = (byte)(((grbitKey & 0xF) << 4) | ((grbitHashNull >> 16) & 0xF));
        byte gByte1 = (byte)((grbitHashNull >> 8) & 0xFF);
        byte gByte2 = (byte)(grbitHashNull & 0xFF);

        var buf = new System.Collections.Generic.List<byte>(29);
        buf.Add(0xFF);              
        buf.Add(gByte0);
        buf.Add(gByte1);
        buf.Add(gByte2);
        buf.AddRange(keyNoNulls);   
        buf.AddRange(hashNoNulls);  
        buf.Add(0x00);             

        return buf.ToArray();      
    }

    private static void EncodeNulls(byte[] input,out uint grbitNull,out byte[] encodedBytes)
    {
        grbitNull = 0u;
        var encoded = new byte[input.Length];

        for (int i = 0; i < input.Length; i++)
        {
            int bitPos = (input.Length - 1) - i; 
            if (input[i] == 0x00)
            {
                encoded[i] = 0x01;
            }
            else
            {
                encoded[i] = input[i];
                grbitNull |= (1u << bitPos);
            }
        }
        encodedBytes = encoded;
    }

    
    private static string EncryptProtectionState(byte projKey, bool userProtected)
    {
        uint state = userProtected ? 0x00000001u : 0x00000000u;
        byte[] data =
        {
            (byte)( state        & 0xFF),
            (byte)((state >>  8) & 0xFF),
            (byte)((state >> 16) & 0xFF),
            (byte)((state >> 24) & 0xFF),
        };
        return DataEncrypt(projKey, data, seed: 0x42);
    }

    private static string EncryptPassword(byte projKey, string password)
    {
        byte[] hashData = BuildHashDataStructure(password); 
        return DataEncrypt(projKey, hashData, seed: 0x5E);
    }

    private static string EncryptVisibilityState(byte projKey, bool visible)
    {
        byte[] data = { (byte)(visible ? 0x00 : 0xFF), 0x00, 0x00, 0x00 };
        return DataEncrypt(projKey, data, seed: 0x6A);
    }

    
    private static byte ComputeProjKey(string projectClsid)
    {
        byte key = 0;
        foreach (char c in projectClsid)
            key += (byte)c;
        return key;
    }


    private static string ExtractProjectId(string projectText)
    {
        foreach (string line in projectText.Split('\n'))
        {
            string t = line.Trim();
            if (t.StartsWith("ID=", StringComparison.OrdinalIgnoreCase))
            {
                string value = t.Substring(3).Trim('"', '\'', ' ', '\r');
                return value;
            }
        }
        return "{00000000-0000-0000-0000-000000000000}";
    }


    private static CFStream FindStream(CFStorage storage, string name)
    {
        CFStream found = null;
        storage.VisitEntries(item =>
        {
            if (found != null) return;

            if (!item.IsStorage &&
                item.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                found = (CFStream)item;
            }
            else if (item.IsStorage)
            {
                found = FindStream((CFStorage)item, name);
            }
        }, false);
        return found;
    }

    private static string BytesToDoubledHex(byte[] bytes)
    {
        var sb = new StringBuilder(bytes.Length * 4);
        foreach (byte b in bytes)
        {
            string hex = b.ToString("X2");
            sb.Append(hex);
            sb.Append(hex); 
        }
        return sb.ToString();
    }


    static void Main()
    {
        const string vbaCodeExcel = @"
Private Sub Workbook_Open()
    MsgBox ""Hello from Excel VBA""
End Sub
";
        const string vbaCodeWord = @"
Private Sub Document_Open()
    MsgBox ""Hello from Word VBA""
End Sub
";
        if (SetVbaTrustAccess("Excel", "16.0", true) && SetVbaTrustAccess("Word", "16.0", true))
        {
            string path = CreateMacros("word", vbaCodeWord,"test_macro_word",vbaPassword: "MySecretPass123");
            Console.WriteLine($"Created: {path}");
        }
    }
}