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

//    [Start]

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

}
