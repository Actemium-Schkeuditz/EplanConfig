/*
Lädt das Menüe um die Unterpunkte zu erzeugen für die Exportfunktionen
mit Prüfung ob Nutzer der BU Schkopau/Schkeuditz angemeldet sind
Christian Langrock
Version 1.0     04.02.2020
update domainName nach Änderung durch IT
V1.0 verschoben auf gitHUB
*/

using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using System;
using System.Security.Principal;
using System.Windows;

namespace EplanCMHL.Ausgabe
{
    class CreateMenue
    {
        readonly string[] domainName = { @"VED\ORG-AC-deDZ002-BU00103-Users", @"VED\ORG-AC-DEMD001-Users", @"VED\ORG-AC-DEMQ002-Users" }; // hier Domain Nammen der Nutzer eintragen


        // Menü zusammenbauen
        [DeclareMenu]
        public void MenuFunction()
        {
            try // Fehlerbehandlung
            {
                if (IsUserInGroup(domainName))
                {
                    //Deklarationen
                    //uint MenuID = new uint(); // Menü-ID vom neu erzeugten Menü   

                    Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
                    // Abfragen der Menue-ID
                    uint MenuID = oMenu.GetPersistentMenuId("Export / Beschriftung...");

                    // MessageBox.Show(MenuID.ToString()); // nur test

                    oMenu.AddMenuItem(
                            "Ausgabe Projektisten", // Name: Menüpunkt
                            "ExcecuteSummaryPartlist", // Name: Action
                            "Ausgabe verschiedener Listen", // Statustext
                            MenuID, // Menü-ID:
                            1, // 1 = Hinter Menüpunkt, 0 = Vor Menüpunkt
                            false, // Seperator davor anzeigen
                            false // Seperator dahinter anzeigen
                        );

                    // Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
                    // uint presMenuId = oMenu.GetPersistentMenuId("Makros"); // nicht verwendet
                    uint presMenuId = 37024; //Menü-ID: Einfügen/Fenstermakro...
                    oMenu.AddMenuItem("Makros Einfuegen mit Navigator",
                        "ShowMacroNavi",
                        "Navigator zum Einfügen von Makros",
                        presMenuId,
                        1,
                        false,
                        false
                        );
                    // Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
                    oMenu.AddMenuItem("PDF (Assistent)...",     // Name: Menüpunkt
                        "PDFAssistent_Start",                   // Name: Action
                        "PDF Assistent," +                      // Statustext
                        " aktuelles Projekt als PDF-Datei exportieren",
                        35287,
                        1,
                        false,
                        false
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

                    //         
                    // string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
                    // loadScripts(scriptName);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [DeclareAction("ExcecuteSummaryPartlist")]

        public void ExcecuteSummaryPartList()
        {
            string scriptName = @"$(MD_SCRIPTS)\Ausgabe\AusgabeListen.cs";
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }


        [DeclareAction("ShowMacroNavi")]

        public void ExcecuteMacroNavigator()
        {
            string scriptName = @"$(MD_SCRIPTS)\MakroNavigator\MacroNaviForm.cs";
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }

        [DeclareAction("PDFAssistent_Start")]

        public void ExcecutePDF_Assistent()
        {
            string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName, "1");
        }



        [DeclareEventHandler("onActionStart.String.XPrjActionProjectClose")]
        public void ExcecuteBackupScript()
        {
            string scriptName = @"$(MD_SCRIPTS)\Backup\BackupOnClosingProject_V03.cs";
            // Pfad auflösen
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }



        private void ExcecuteScripts(string sScriptName, string Param = "")
        {
            try
            {
                if (IsUserInGroup(domainName))
                {
                    // Ausgabe der Einbauorte 
                    // Parameter für die Action
                    ActionCallingContext actionCallingContext = new ActionCallingContext();
                    actionCallingContext.AddParameter("ScriptFile", sScriptName);
                    if (!string.IsNullOrEmpty(Param))
                    {
                        actionCallingContext.AddParameter("Param1", Param);
                    }
                    // ausführen
                    new CommandLineInterpreter().Execute("ExecuteScript", actionCallingContext);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }

        }

        


        private void loadScripts(string sScriptName)
        {
            try
            {
                if (IsUserInGroup(domainName))
                {
                    // Parameter für die Action
                    ActionCallingContext actionCallingContext = new ActionCallingContext();
                    actionCallingContext.AddParameter("ScriptFile", sScriptName);
                    actionCallingContext.AddParameter("RegisterAgain", "1");        // Ermöglicht, das Scipt erneut zu registrieren
                    actionCallingContext.AddParameter("ShowDecider", "0");          // Blendet eine Abfrage ein, ob das Scipt erneut registriert werden soll 

                    // ausführen
                    new CommandLineInterpreter().Execute("RegisterScript", actionCallingContext);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
           
        }


        public static bool IsUserInGroup(string[] gruopName)
        {
            if (string.IsNullOrWhiteSpace(gruopName[0]))
            {
                throw new ArgumentException("Fehler bei Benutztergruppe");
               
            }

            ///<summary>
            /// Prüfen ob der angemeldete User in der richtigen Windowsdomain ist
            /// </summary>

            ///<returns>
            /// 1 entspricht der Nutzer ist in der Windowsdomain
            /// </returns>

            try
            {
                foreach (string grName in gruopName)
                    {
                    foreach (IdentityReference group in WindowsIdentity.GetCurrent().Groups)
                    {
                        if (grName == group.Translate(typeof(NTAccount)).Value)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
        }


        // ab hier alles für PDF Assistent
        [DeclareEventHandler("onActionStart.String.XPrjActionProjectClose")]
        public void ExcecutPDFonCloseProject()
        {
            if (IsUserInGroup(domainName))
            {
                //Einstellung einlesen
                Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
                if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose"))
                {
                    bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose", 1);
                    if (bChecked) //Bei ProjectClose ausführen
                    {
                        PDFAssistent_SollStarten();
                    }
                }
                return;
            }
        }

        [DeclareEventHandler("Eplan.EplApi.OnMainEnd")]
        public void ExcecutePDFonEplanEnd()
        {
            if (IsUserInGroup(domainName))
            {
                //Einstellung einlesen
                Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
                if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd"))
                {
                    bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd", 1);
                    if (bChecked) //Bei EplanEnd ausführen
                    {
                        PDFAssistent_SollStarten();
                    }
                }
                return;
            }
        }

        public void PDFAssistent_SollStarten()
        {
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

            //Ist Projekt in Projekt-Ordner
            //Wenn angekreuzt dann muß Projekt im Ordner sein für Assistent, sonst kein Assistent
            //Wenn nicht angekreuzt dann Assistent
            if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage"))
            {
                bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner", 1);
                string sProjektInOrdner = oSettings.GetStringSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName", 0);
                if (bChecked)
                {
                    string sProjektOrdner = PathMap.SubstitutePath("$(PROJECTPATH)");
                    sProjektOrdner = sProjektOrdner.Substring(0, sProjektOrdner.LastIndexOf(@"\"));
                    if (!sProjektOrdner.EndsWith(@"\"))
                    {
                        sProjektOrdner += @"\";
                    }
                    if (sProjektInOrdner == sProjektOrdner) //hier vielleicht noch erweitern auf alle Unterordner (InString?)
                    {
                        PDFAssistent_ausführen();
                    }
                    else
                    {
                       // Close();
                    }
                }
                else
                {
                    PDFAssistent_ausführen();
                }
            }
        }

        public void PDFAssistent_ausführen()
        {
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
            if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage"))
            {
                bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage", 1);
                if (bChecked)
                {
                    string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
                    scriptName = PathMap.SubstitutePath(scriptName);
                    ExcecuteScripts(scriptName, "0");
                    //  cboAusgabeNach.SelectedIndex = 0;
                    //  ReadSettings();
                    //  PDFexport(txtSpeicherort.Text + txtDateiname.Text + @".pdf");
                    //  Close();
                }
                else
                {
                    //PDFAssistent_Start();
                    string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
                    scriptName = PathMap.SubstitutePath(scriptName);
                    ExcecuteScripts(scriptName, "1");
                }
            }
        }

    }
}