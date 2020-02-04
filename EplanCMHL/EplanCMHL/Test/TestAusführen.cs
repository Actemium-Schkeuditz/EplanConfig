// Ausführen über die Option in Symbolleiste: ExecuteScript /ScriptFile:"$(MD_SCRIPTS)\TestAusführen.cs"

using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;
using System.Windows;
using System;
using System.IO;
//using IWshRuntimeLibrary;


public class SimpleScript
{

    [Start]
    public void MyFunction()
    {

        // MessageBox.Show("MyFunction was called!", "SimpleScript");


        // Spielen mit Import und Export Schematas
        // prüfen ob die das Schema vorhanden ist
        string schemeName = "ExportFunc_Einbauorte";
        string SETTINGS_PATH = "USER.Labelling.Config";
        string sImportFile = @"$(MD_SCRIPTS)\Test\LB.ExportFunc_Einbauorte.xml";

        // EplanMacro mit Pfad zum Script Ordner auflösen
        string saufgeloesst = PathMap.SubstitutePath(sImportFile);
        //MessageBox.Show(saufgeloesst);

        // Import eines bestimmten Schemas
        bool result = ImportScheme(schemeName, SETTINGS_PATH, saufgeloesst);
        if (result == false)
        {
            MessageBox.Show("Fehler beim Import");
        }
        return;
    }

    private static bool ImportScheme (string schemeName, string SETTINGS_PATH, string sFilename)
    {
        /// <summary>
        /// Import eines Eplan Schemes
        /// </summary>
        /// <remarks>
        /// string schemeName Name des Schematas
        /// string SETTINGS_PATH = wenn unklar Daten aus XML lesen Beispiel für Beschriftung: "USER.Labelling.Config";
        /// string sFilname Verweis auf XML-Datei, Beispiel: @"$(MD_SCRIPTS)\Test\LB.Einbauorte.xml"
        /// </remarks>

        SchemeSetting schemeSetting = new SchemeSetting();
        schemeSetting.Init(SETTINGS_PATH);
        try
        {
            // prüfen ob die das Schema vorhanden ist, import wenn nicht da
            if (schemeSetting.CheckIfSchemeExists(schemeName) == false)
            {
                // testen ob es die XML Datei gibt
                if (FunctionCheckFileExist(sFilename))
                {
                    // import wenn Schema nicht in Eplan und XML-Datei existiert
                    schemeSetting.ImportSchemes(sFilename, false);
                    return true;
                }
                else
                {
                    MessageBox.Show("XML-Datei fehlt: " + sFilename);
                    return false;
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            return false;
        }
    }

    private static bool FunctionCheckFileExist(string sfileName)
        {
        /// <summary>
        /// Testen ob eine Datei existiert
        /// <\summary>
        /// <remarks>
        /// return bool 1 = file exist
        /// </remarks>
        try
        {
            if (System.IO.File.Exists(sfileName))
            {
              //  MessageBox.Show("Datei vorhanden.");
                return true;
            }
            else
            {
               // MessageBox.Show("Datei fehlt.");
                return false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            throw;
        }

    }
  /*
    private static void CreateShortcut()
    {

        string filname = @"C:\Daten\Eplan\Datesicherung_SK001_25A_01.xlsx";
        string workingDir = @"C:\Daten\Eplan\";
        var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var shell = new WshShell();
        var shortCutLinkFilePath = Path.Combine(startupFolderPath, @"\CreateShortcutSample.lnk");
        var windowsApplicationShortcut = (IWshShortcut)shell.CreateShortcut(shortCutLinkFilePath);
        windowsApplicationShortcut.Description = "How to create short for application example";
        windowsApplicationShortcut.WorkingDirectory = workingDir;
        windowsApplicationShortcut.TargetPath = filname;
        windowsApplicationShortcut.Save();
    }
    private void MakeShortcut(string shortcutFileName, string targetFileName)
    {
        IWshShell shell = new WshShell();
        IWshShortcut shortcut = shell.CreateShortcut(shortcutFileName) as IWshShortcut;
        shortcut.TargetPath = targetFileName;
        shortcut.Description = "Made with WinR!!! \n" + targetFileName;//Add path
        shortcut.WorkingDirectory = new FileInfo(targetFileName).Directory.FullName;
        shortcut.Save();
    }
    */
}
