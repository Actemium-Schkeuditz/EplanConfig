using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System;

public class AusgabeEinbauorte
{
    [DeclareAction("MenueAusgabeEinbauOrte")]



    public void Function()
    {
        string strProjectName = GetProject();
        string strLabelName = PathMap.SubstitutePath("$(PROJECTNAME)");
        string myDate = System.DateTime.Now.ToString("yyyyMMdd");
        string myTime = System.DateTime.Now.ToString("HHmm");


        // Dokumentenordner auswählen
        string sProjectDocPath = PathMap.SubstitutePath("$(DOC)");
        string sfilenameEinbauort = sProjectDocPath + @"\Einbauorte_" + myDate + "_" + myTime + ".txt";

        // Ausgabe der TXT mit Einbauorten
        LabellingEinbauort(sfilenameEinbauort, strProjectName);

        // Textdatei lesen
        List<string> sEinbauort = ReadFile(sfilenameEinbauort);
        // Zeilenweise zerlegen
        foreach (var item in sEinbauort)
        {
            //  MessageBox.Show(item);
            ExportPartlist(sProjectDocPath, strLabelName, strProjectName,item,myDate, myTime);
                }

        string sExportierteOrte = string.Join(",\r\n+", sEinbauort.ToArray());

        MessageBox.Show("Artikeldaten exportiert für:\r\n" + sExportierteOrte);

    }

    // Menü zusammenbauen
    [DeclareMenu]
    public void MenuFunction()
    {
        //Deklarationen
        //uint MenuID = new uint(); // Menü-ID vom neu erzeugten Menü   

        Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
        // Abfragen der Menue-ID
        uint MenuID = oMenu.GetPersistentMenuId("Export / Beschriftung...");

        // MessageBox.Show(MenuID.ToString()); // nur test

        oMenu.AddMenuItem(
                "Stückliste nach Einbauort", // Name: Menüpunkt
                "MenueAusgabeEinbauOrte", // Name: Action
                "Ausgabe der Einbauorte als XML", // Statustext
                MenuID, // Menü-ID:
                1, // 1 = Hinter Menüpunkt, 0 = Vor Menüpunkt
                false, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
            );
        // Hauptmenü mit einem Unterpunkt
        /*  MenuID = oMenu.AddMainMenu(
                     "Ausgabe", // Name: Menü
                  "Hilfe", // neben...
                  "Info", // Name: Menüpunkt
                 "ActionInfo", // Name: Action
                  "Info Einstellungen", // Statustext
                  1 // 1 = Hinter Menüpunkt, 0 = Vor Menüpunkt
                  );
          // Menüpunkt Übersetzungsdatenbank
          MenuIDTrans = oMenu.AddPopupMenuItem(
                  "Ausgabe Excel", // Name: Menü
                  "Einbauorte", // Name: Menüpunkt
                  "MenueAusgabeEinbauOrte", // Name: Action
                  "Ausgabe der Einbauorte als XML", // Statustext
                  MenuID, // Menü-ID:
                  1, // 1 = Hinter Menüpunkt, 0 = Vor Menüpunkt
                  false, // Seperator davor anzeigen
                  false // Seperator dahinter anzeigen
                  );
              */
    }

    private string GetProject()
    {
        // Rückgabe des Projektnammens mit Erweiterung
        string strProject = "";
        ActionCallingContext ProjectContext = new ActionCallingContext();
        ProjectContext.AddParameter("TYPE", "PROJECT");

        new CommandLineInterpreter().Execute("selectionset", ProjectContext);
        ProjectContext.GetParameter("PROJECT", ref strProject);
        return strProject;

    }

    public List<string> ReadFile(string fileName)
    {
        //Rückgabe als Liste
        List<String> lines = new List<String>();

        using (StreamReader sr = new StreamReader(fileName, Encoding.Default))
        {
            while (sr.Peek() != -1)
                lines.Add(sr.ReadLine());
        }
        return lines;
    }
     private static void ExportPartlist(string sProjectDocPath, string strLabelName, string strProjectName, string sLoc, string myDate, string myTime)
        {
       
            using (Settings oSett = new Settings())
            {
            oSett.SetStringSetting("USER.Labelling.Config.Summarized parts list.Data.SortFilter.FilterSchemeData", "0|1|0|1420;0|0||0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sLoc + "|0|1|1|0|0|0;0|", 0);
            // Filter für Aufstellungsort undEinbauort 

        }

            // Parameter für die Action
            ActionCallingContext labellingContext = new ActionCallingContext();

            labellingContext.AddParameter("PROJECTNAME", strProjectName); // Projektname komplett mit Erweiterung
            labellingContext.AddParameter("CONFIGSCHEME", "Summarized parts list");
            labellingContext.AddParameter("FILTERSCHEME", ""); // Hier darf kein Filterschema eingetragen werden
            labellingContext.AddParameter("SORTSCHEME", "");
            labellingContext.AddParameter("LANGUAGE", "de_DE");
            labellingContext.AddParameter("DESTINATIONFILE", sProjectDocPath + @"\Artikelsummen" + "_" + strLabelName + "_" + sLoc + "_" + myDate + "_" + myTime + ".xlsx"); // Zieldatei: 
            labellingContext.AddParameter("RECREPEAT", "1");
            labellingContext.AddParameter("TASKREPEAT", "1");
            labellingContext.AddParameter("SHOWOUTPUT", "0");

            // Ausführen der Beschriftungsaction mit Parametern
            new CommandLineInterpreter().Execute("label", labellingContext);
        }

private static void LabellingEinbauort(string sfilename, string strProjectName)
    {
        // Ausgabe der Einbauorte 
        // Parameter für die Action
        ActionCallingContext labellingContext = new ActionCallingContext();
        labellingContext.AddParameter("PROJECTNAME", strProjectName); // Projektname komplett mit Erweiterung
        labellingContext.AddParameter("CONFIGSCHEME", "Einbauorte");
        labellingContext.AddParameter("FILTERSCHEME", ""); // Hier darf kein Filterschema eingetragen werden
        labellingContext.AddParameter("SORTSCHEME", "");
        labellingContext.AddParameter("LANGUAGE", "de_DE");
        labellingContext.AddParameter("DESTINATIONFILE", sfilename); // Zieldatei: "Projektpfad/Projektname_Testbeschriftung.xls"
        labellingContext.AddParameter("RECREPEAT", "1");
        labellingContext.AddParameter("TASKREPEAT", "1");
        labellingContext.AddParameter("SHOWOUTPUT", "0");
        labellingContext.AddParameter("USESELECTION", "0");   //auswahlh berücksichtigen: 0 = Nein, 1 = Ja

        // Ausführen der Beschriftungsaction mit Parametern
        new CommandLineInterpreter().Execute("label", labellingContext);
    }

}

