using System;
using System.IO;
// Add reference to COM library: "Windows Script Host Object Model" (wshom.ocx)
// In Visual Studio: Right-click project > Add > Reference... > COM > "Windows Script Host Object Model"
using IWshRuntimeLibrary;

namespace FisiereContorPagini.InstallerHelpers
{
    public static class ShortcutHelper
    {
        // Creates a .lnk on the desktop pointing to exePath.
        // If forAllUsers is true, attempts the common (All Users) desktop first and falls back to current user desktop on permission failure.
        public static void CreateDesktopShortcut(string exePath, string shortcutName = "FisiereContorPagini", bool forAllUsers = false)
        {
            if (string.IsNullOrEmpty(exePath) || !System.IO.File.Exists(exePath))
                throw new FileNotFoundException(nameof(exePath));

            string desktop;
            if (forAllUsers)
            {
                // Try common desktop (requires elevated privileges)
                desktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
            }
            else
            {
                desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            }

            string shortcutPath = Path.Combine(desktop, shortcutName + ".lnk");

            try
            {
                var shell = new WshShell();
                var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = exePath;
                shortcut.WorkingDirectory = Path.GetDirectoryName(exePath);
                shortcut.WindowStyle = 1;
                shortcut.Description = "FisiereContorPagini";
                // Use the executable's icon
                shortcut.IconLocation = exePath + ",0";
                shortcut.Save();
            }
            catch (UnauthorizedAccessException) when (forAllUsers)
            {
                // If we failed to write to the common desktop due to permissions, fall back to current user's desktop
                string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fallbackPath = Path.Combine(userDesktop, shortcutName + ".lnk");

                var shell = new WshShell();
                var shortcut = (IWshShortcut)shell.CreateShortcut(fallbackPath);
                shortcut.TargetPath = exePath;
                shortcut.WorkingDirectory = Path.GetDirectoryName(exePath);
                shortcut.WindowStyle = 1;
                shortcut.Description = "FisiereContorPagini";
                shortcut.IconLocation = exePath + ",0";
                shortcut.Save();
            }
        }
    }
}