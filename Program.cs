using System.Reflection;
using IWshRuntimeLibrary;

namespace LolibarAutorun;
class StartUpScript {
    public static void CreateShortcut(string targetPath, string shortcutPath)
    {
        WshShell shell = new();
        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

        shortcut.TargetPath = targetPath;
        shortcut.IconLocation = targetPath;

        shortcut.Save();
    }
    public static void Main(string[] args) {
        var orig_path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var path = orig_path?.Replace(orig_path.Split('\\').Last(), "");
        
        if (Path.Exists("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\lolibar.lnk"))
        {
            System.IO.File.Delete("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\lolibar.lnk");
        }
        else
        {
            try
            {
                CreateShortcut(
                    $"{path}lolibar.exe",
                    "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\lolibar.lnk"
                );
            }
            catch (Exception e)
            {
                Console.WriteLine(@"A problem to create Lolibar's shortcut in: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup...");
                Console.WriteLine($"Execution path: {orig_path}.");
                Console.WriteLine($"Expected lolibar.exe's path: {path}lolibar.exe.");
                Console.WriteLine();
                Console.WriteLine(e);
                Console.ReadKey();
            }
        }
    }
}