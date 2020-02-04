/*
	NAME....: MenuDatenquelle
	USAGE...: for EPLAN P8 (v2.5)
	AUTHOR..: C. Langrock
	VERSION.: 2016-07-18
	FUNC....: Umschalten zwischen den Datenquellen f�r Artikel, W�rterbuch und Projektdatenbank; Umschalten der Verzeichnisse
*/
using Eplan.EplApi.Scripting;

public class MenueDaten
{   
    // Datenquellen einstellen
    // Pfad zur XML Datei + Dateiname
    // Verzeichnisse
    string XMLFileVerzServer = "$(EPLAN_DATA)" + "\\Skripte\\CMHL\\Auswahl_Datenquelle\\Verzeichnisse_Server.xml";
    // Achtung!!! auch  "$(EPLAN_DATA)" + "\\Scripte\\... m�glich je nach instalation
    // immer relative zum Standardverzeichnis aufrufen, nicht mit �(EPLAN_SCRIPT) 

    // SQL Server und Instanzen
    string sSqlLocal = "CMSK001NBD706\\CMHL_EPLAN";
    string sSqlFern = "EPLAN01P\\EPLAN01P";

    // Datenbanken Server
    string sPartsDBFern = "SK001_25A_01";
    string sPojectDBFern = "SK001_25P_01";
    string sTranslationDBFern = "SK001_25W_01";

    // Datenbanken Lokal
    string sPartsDBLocal = "CMHL_Artikel";
    string sPojectDBLocal = "CMHL_PROJEKT";
    string sTranslationDBLocal = "CMHL_Woerterbuch";

    [DeclareAction("MenuActionPartsLocal")]
    public void ActionFunction_PartsLocal()
    {
        //Strings zusammenbauen
        string sSqlDB = sSqlLocal + "\n" + sPartsDBLocal + "\n0\n\n\n1";

        // DB set direkt
        ActionCallingContext SqlSet = new ActionCallingContext();
        SqlSet.AddParameter("set", "USER.PartsManagementGui.Database");
        SqlSet.AddParameter("value", sSqlDB);
        new CommandLineInterpreter().Execute("XAfActionSetting", SqlSet);

        // R�ckmeldung
        MessageBox.Show("Lokale Daten werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    [DeclareAction("MenuActionPrjLocal")]
    public void ActionFunction_PrjLocal()
    {
        //Strings zusammenbauen
        string sSqlDB = sSqlLocal + "\n" + sPojectDBLocal + "\n0\n\n\n1";

        // DB set direkt
        ActionCallingContext SqlSet = new ActionCallingContext();
        SqlSet.AddParameter("set", "COMPANY.PrjManagementGUI.Database.SQL");
        SqlSet.AddParameter("value", sSqlDB);
        new CommandLineInterpreter().Execute("XAfActionSetting", SqlSet);
        // R�ckmeldung
        MessageBox.Show("Lokale Daten werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    [DeclareAction("MenuActionTransLocal")]
    public void ActionFunction_TransLocal()
    {
        //Strings zusammenbauen
        string sSqlDB = sSqlLocal + "\n" + sTranslationDBLocal + "\n0\n\n\n1";

        // DB set direkt
        ActionCallingContext SqlSet = new ActionCallingContext();
        SqlSet.AddParameter("set", "USER.TRANSLATEGUI.DATABASE_NAME_SQL");
        SqlSet.AddParameter("value", sSqlDB);
        new CommandLineInterpreter().Execute("XAfActionSetting", SqlSet);
        // R�ckmeldung
        MessageBox.Show("Lokale Daten werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    [DeclareAction("MenuActionPartsServer")]
    public void ActionFunction_PartsServer()
    {
        //Strings zusammenbauen
        string sPartsSqlDB = sSqlFern + "\n" + sPartsDBFern + "\n0\n\n\n1";

        // DB set direkt
        ActionCallingContext SqlSet = new ActionCallingContext();
        SqlSet.AddParameter("set", "USER.PartsManagementGui.Database");
        SqlSet.AddParameter("value", sPartsSqlDB);
        new CommandLineInterpreter().Execute("XAfActionSetting", SqlSet);

         // R�ckmeldung
        MessageBox.Show("Server Daten werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    [DeclareAction("MenuActionPrjServer")]
    public void ActionFunction_PrjServer()
    {
        //Strings zusammenbauen
        string sSqlDB = sSqlFern + "\n" + sPojectDBFern + "\n0\n\n\n1";

        // DB set direkt
        ActionCallingContext SqlSet = new ActionCallingContext();
        SqlSet.AddParameter("set", "COMPANY.PrjManagementGUI.Database.SQL");
        SqlSet.AddParameter("value", sSqlDB);
        new CommandLineInterpreter().Execute("XAfActionSetting", SqlSet);
        // R�ckmeldung
        MessageBox.Show("Server Daten werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    [DeclareAction("MenuActionTransServer")]
    public void ActionFunction_TransServer()
    {
        //Strings zusammenbauen
        string sSqlDB = sSqlFern + "\n" + sTranslationDBFern + "\n0\n\n\n1";

        // DB set direkt
        ActionCallingContext SqlSet = new ActionCallingContext();
        SqlSet.AddParameter("set", "USER.TRANSLATEGUI.DATABASE_NAME_SQL");
        SqlSet.AddParameter("value", sSqlDB);
        new CommandLineInterpreter().Execute("XAfActionSetting", SqlSet);
        // R�ckmeldung
        MessageBox.Show("Server W�rterbuch wird genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    // Verzeichnisse
    [DeclareAction("MenuActionVerServer")]
    public void ActionFunction_VerServer()
    {

	// Verzeichnisse wechseln	   
	ActionCallingContext VerzChange = new ActionCallingContext();

	// Vorbereiten - Einstellungen �ber die Aktion "XSettingsImport" 
	VerzChange.AddParameter("XMLFile", XMLFileVerzServer);

	// Jetzt Einstellungen �ber die Aktion "XSettingsImport" importieren
	new CommandLineInterpreter().Execute("XSettingsImport", VerzChange);

	ActionCallingContext VerzSet = new ActionCallingContext();
    // falls Import nicht reicht setzen
	VerzSet.AddParameter("set", "USER.ModalDialogs.PathsScheme.LastUsed");
	VerzSet.AddParameter("value", "Server_GK");
 	new CommandLineInterpreter().Execute("XAfActionSetting", VerzSet);

// R�ckmeldung
    MessageBox.Show("Verzeichnisse auf dem Server werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    [DeclareAction("MenuActionVerLokal")]
    public void ActionFunction_VerLokal()
    {
// kein Import nur wechseln
	ActionCallingContext VerzSet = new ActionCallingContext();

	VerzSet.AddParameter("set", "USER.ModalDialogs.PathsScheme.LastUsed");
	VerzSet.AddParameter("value", "Standard");

 	new CommandLineInterpreter().Execute("XAfActionSetting", VerzSet);

// R�ckmeldung
    MessageBox.Show("Lokale Verzeichnisse werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }


    [DeclareAction("MenuActionAllLocal")]
    public void ActionFunction_AllLocal()
    {
        ActionFunction_PartsLocal();
        ActionFunction_TransLocal();
        ActionFunction_VerLokal();
        ActionFunction_PrjLocal();

    // R�ckmeldung
    //   MessageBox.Show("Lokale Verzeichnisse und Datenbanken werden genutzt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

        return;
    }

    // Info
    [DeclareAction("ActionInfo")]
    public void ActionFunctionInfo()
    {       
        Eplan.EplApi.Base.Settings oSettingsRead = new Eplan.EplApi.Base.Settings();
        // Datenbank Einstellung auslesen
        string sArtikelDB =  oSettingsRead.GetStringSetting("USER.PartsManagementGui.Database", 0);
        string sProjektDB = oSettingsRead.GetStringSetting("COMPANY.PrjManagementGUI.Database.SQL", 0);
        string sTranslateDB = oSettingsRead.GetStringSetting("USER.TRANSLATEGUI.DATABASE_NAME_SQL", 0);

        // Verzeichniseinstellungen auslesen
        string sVerzeichnis = oSettingsRead.GetStringSetting("USER.ModalDialogs.PathsScheme.LastUsed", 0); // + SchemeName + ".Data.", 0);
       
        // �berfl�ssige Zeichen abschneiden
        int len = sArtikelDB.Length > 0 ? sArtikelDB.Length - 6 : 0;
        sArtikelDB = sArtikelDB.Substring(0, len);

        len = sProjektDB.Length > 0 ? sProjektDB.Length - 6 : 0;
        sProjektDB = sProjektDB.Substring(0, len);

        len = sTranslateDB.Length > 0 ? sTranslateDB.Length - 6 : 0;
        sTranslateDB = sTranslateDB.Substring(0, len);
        // Ausgabe
        MessageBox.Show("Artikeldatenbank: " + sArtikelDB + Environment.NewLine + Environment.NewLine
          +  "Verzeichnis: " + sVerzeichnis + Environment.NewLine + Environment.NewLine
          + "Projektdatenbank: " + sProjektDB + Environment.NewLine + Environment.NewLine
          + "Uebersetzungsdatenbank: " + sTranslateDB, "Aktuelle Einstellungen", MessageBoxButtons.OK, MessageBoxIcon.Information);
      
        return;
    }

    // Men� zusammenbauen
    [DeclareMenu]
    public void MenuFunction()
    {
       	//Deklarationen
        uint MenuID = new uint(); // Men�-ID vom neu erzeugten Men�   
	    uint MenuIDSQL = new uint(); // Men�-ID vom neu erzeugten untermenen�
	    uint MenuIDVer = new uint(); // Men�-ID vom neu erzeugten untermenen�
        uint MenuIDPrj = new uint(); // Men�-ID vom neu erzeugten untermenen�
        uint MenuIDTrans = new uint(); // Men�-ID vom neu erzeugten untermenen�

        Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();

        // Hauptmen� mit einem Unterpunkt
        MenuID = oMenu.AddMainMenu(
               	"Datenquelle", // Name: Men�
                "Hilfe", // neben...
               	"Info", // Name: Men�punkt
               "ActionInfo", // Name: Action
                "Info Einstellungen", // Statustext
                1 // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                );

        // Men�punkt �bersetzungsdatenbank
        MenuIDTrans = oMenu.AddPopupMenuItem(
                "Uebersetzungsdatenbank", // Name: Men�
                "Lokale Daten", // Name: Men�punkt
                "MenuActionTransLocal", // Name: Action
                "Auswahl lokale �bersetztungsdatenbank", // Statustext
                MenuID, // Men�-ID:
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                false, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
                );

        //Serverdaten
        oMenu.AddMenuItem(
                "Serverdaten", // Name: Men�punkt 
                "MenuActionTransServer", // Name: Action
                "Auswahl �bersetzungsdatenbank vom Server", // Statustext
                MenuIDTrans, // Men�-ID: 
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                false, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
            );

        // Men�punkt Projekt Datenbank
        MenuIDPrj = oMenu.AddPopupMenuItem(
                "Projektdatenbank", // Name: Men�
                "Lokale Daten", // Name: Men�punkt
                "MenuActionPrjLocal", // Name: Action
                "Auswahl lokale Projektdatenbank", // Statustext
                MenuID, // Men�-ID:
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                false, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
                );

        //Serverdaten
        oMenu.AddMenuItem(
                "Serverdaten", // Name: Men�punkt 
                "MenuActionPrjServer", // Name: Action
                "Auswahl Projektdatenbank vom Server", // Statustext
                MenuIDPrj, // Men�-ID: 
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                false, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
            );

        // Men�punkt Artikel
        MenuIDSQL = oMenu.AddPopupMenuItem(
                "Artikel Daten", // Name: Men�
                "Lokale Daten", // Name: Men�punkt
                "MenuActionPartslocal", // Name: Action
                "Auswahl lokale Daten", // Statustext
                MenuID, // Men�-ID: 
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                true, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
                );

        oMenu.AddMenuItem(
            	"Serverdaten", // Name: Men�punkt 
            	"MenuActionPartsServer", // Name: Action
		        "Auswahl SQL Daten vom Server", // Statustext
 		        MenuIDSQL, // Men�-ID: 
		        1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
            	false, // Seperator davor anzeigen
            	false // Seperator dahinter anzeigen
            );

        // Men�punkt Verzeichnisse
	    MenuIDVer = oMenu.AddPopupMenuItem(
                "Verzeichnisse", // Name: Men�
                "Lokale Ordner", // Name: Men�punkt
                "MenuActionVerLokal", // Name: Action
                "Auswahl lokale Verzeichnisse", // Statustext
                MenuID, // Men�-ID:
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                true, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
                );

        //Serverdaten
        oMenu.AddMenuItem(
            	"Server Ordner", // Name: Men�punkt 
            	"MenuActionVerServer", // Name: Action
		        "Auswahl SQL Daten vom Server", // Statustext
 		        MenuIDVer, // Men�-ID: 
		        1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
            	false, // Seperator davor anzeigen
            	false // Seperator dahinter anzeigen
            );

        // alles lokal
        oMenu.AddMenuItem(
                "alles Lokal", // Name: Men�punkt 
                "MenuActionAllLocal", // Name: Action
                "Alle Daten und Verzeichnisse vom PC", // Statustext
                MenuID, // Men�-ID: 
                1, // 1 = Hinter Men�punkt, 0 = Vor Men�punkt
                false, // Seperator davor anzeigen
                false // Seperator dahinter anzeigen
            );

        return;
    }
}