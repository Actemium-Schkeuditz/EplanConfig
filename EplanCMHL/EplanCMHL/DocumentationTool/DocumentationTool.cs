// Documentation-Tool.cs
//
// Aufruf über einen neuen Menüpunkt "Lesezeichen"
// im Untermenü "Projekt"
//
// Copyright by Frank Schöneck, 2017
//
// letzte Änderung:
// V2.0.0, 08.08.2017, Frank Schöneck,	Projektbeginn
// V2.1.0, 17.08.2017, Frank Schöneck,	Erweitert auf 20 Externe Dokumente bei Artikel (eingelagert)
// V2.2.0, 18.08.2017, Frank Schöneck,	Kontextmenüpunkte in Zielverzeichnis hinzugefügt
// V2.3.0, 25.08.2017, Frank Schöneck,	Beschreibung des Dokumentes, der Hersteller und die Artikelnummer wird jetzt angezeigt
//										Kontextmenü in Dokumentenliste hinzugefügt, darüber ist nun "Dokument öffnen" möglich
//										Kontextmenü Zielverzeichnis um Menüpunkt "Hersteller zur Verzeichnisstruktur hinzufügen" erweitert
//										Kontextmenü Zielverzeichnis um Menüpunkt "Artikelnummer zur Verzeichnisstruktur hinzufügen" erweitert
// V2.4.0, 06.09.2017, Frank Schöneck,	Die Oberfläche ist nun Größenveränderbar und die Position und Größe wird gespeichert
//										Button "Extras" hinzugefügt, zum Aufrufen des Kontextmenü Zielverzeichnis
//										Die Einstellungen des des Kontextmenü Zielverzeichnis werden nun gespeichert
//										Die Spalten können nun durch anklicken der Spaltenüberschrift sortiert werden
// V2.4.1, 19.12.2017, Frank Schöneck,	Fehlerbehandlung für Dateikopieren eingefügt.
//										Beim Zielverzeichnisnamen werden ungültige Zeichen automatisch entfernt
//
// V2.4.1.ACT	16.06.2021 Christian Langrock,	Integration in VINCI Citrix	
// für Eplan Electric P8, V2.6
//

/*
The following compiler directive is necessary to enable editing scripts
within Visual Studio.
It requires that the "Conditional compilation symbol" SCRIPTENV be defined 
in the Visual Studio project properties
This is because EPLAN's internal scripting engine already adds "using directives"
when you load the script in EPLAN. Having them twice would cause errors.
*/

#if SCRIPTENV
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
#endif

/*
On the other hand, some namespaces are not automatically added by EPLAN when
you load a script. Those have to be outside of the previous conditional compiler directive
*/

using System;
using System.IO;
using System.Xml;
using System.Collections;
using System.Windows.Forms;


public partial class frmDocumentationTool : System.Windows.Forms.Form
{
	private Button btnAbbrechen;
	private ListView listView;
	private ColumnHeader columnFileName;
	private Button btnKopieren;
	private TextBox txtZielverzeichnis;
	private Label label1;
	private Button btnOrdnerWählen;
	private Button btnOdnerÖffnen;
	private ContextMenuStrip contextMenuZielverzeichnis;
	private ToolStripMenuItem toolStripMenuItemProjektVerzeichnisstruktur;
	private ToolStripMenuItem toolStripMenuItemProjekteVerzeichnis;
	private ColumnHeader columnFileDescription;
	private ContextMenuStrip contextMenuListView;
	private ToolStripMenuItem toolStripMenuItemDocumentOpen;
	private ColumnHeader columnHersteller;
	private ToolStripMenuItem toolStripMenuHerstellerVerzeichnis;
	private ToolStripSeparator toolStripSeparator1;
	private ColumnHeader columnArtikelnummer;
	private ToolStripMenuItem toolStripMenuArtikelnummerVerzeichnis;
	private Button btnExtras;

	#region Vom Windows Form-Designer generierter Code

	/// <summary>
	/// Erforderliche Designervariable.
	/// </summary>
	private System.ComponentModel.IContainer components = null;

	/// <summary>
	/// Verwendete Ressourcen bereinigen.
	/// </summary>
	/// <param name="disposing">True, wenn verwaltete Ressourcen
	/// gelöscht werden sollen; andernfalls False.</param>
	protected override void Dispose(bool disposing)
	{
		if (disposing && (components != null))
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	/// <summary>
	/// Erforderliche Methode für die Designerunterstützung.
	/// Der Inhalt der Methode darf nicht mit dem Code-Editor
	/// geändert werden.
	/// </summary>
	private void InitializeComponent()
	{
		this.components = new System.ComponentModel.Container();
		this.btnAbbrechen = new System.Windows.Forms.Button();
		this.listView = new System.Windows.Forms.ListView();
		this.columnFileName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
		this.columnFileDescription = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
		this.columnHersteller = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
		this.columnArtikelnummer = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
		this.contextMenuListView = new System.Windows.Forms.ContextMenuStrip(this.components);
		this.toolStripMenuItemDocumentOpen = new System.Windows.Forms.ToolStripMenuItem();
		this.btnKopieren = new System.Windows.Forms.Button();
		this.txtZielverzeichnis = new System.Windows.Forms.TextBox();
		this.contextMenuZielverzeichnis = new System.Windows.Forms.ContextMenuStrip(this.components);
		this.toolStripMenuHerstellerVerzeichnis = new System.Windows.Forms.ToolStripMenuItem();
		this.toolStripMenuArtikelnummerVerzeichnis = new System.Windows.Forms.ToolStripMenuItem();
		this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
		this.toolStripMenuItemProjekteVerzeichnis = new System.Windows.Forms.ToolStripMenuItem();
		this.toolStripMenuItemProjektVerzeichnisstruktur = new System.Windows.Forms.ToolStripMenuItem();
		this.label1 = new System.Windows.Forms.Label();
		this.btnOrdnerWählen = new System.Windows.Forms.Button();
		this.btnOdnerÖffnen = new System.Windows.Forms.Button();
		this.btnExtras = new System.Windows.Forms.Button();
		this.contextMenuListView.SuspendLayout();
		this.contextMenuZielverzeichnis.SuspendLayout();
		this.SuspendLayout();
		// 
		// btnAbbrechen
		// 
		this.btnAbbrechen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnAbbrechen.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.btnAbbrechen.Location = new System.Drawing.Point(594, 393);
		this.btnAbbrechen.Name = "btnAbbrechen";
		this.btnAbbrechen.Size = new System.Drawing.Size(120, 26);
		this.btnAbbrechen.TabIndex = 6;
		this.btnAbbrechen.Text = "Abbrechen";
		this.btnAbbrechen.UseVisualStyleBackColor = true;
		this.btnAbbrechen.Click += new System.EventHandler(this.btnAbbrechen_Click);
		// 
		// listView
		// 
		this.listView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
		| System.Windows.Forms.AnchorStyles.Left)
		| System.Windows.Forms.AnchorStyles.Right)));
		this.listView.AutoArrange = false;
		this.listView.CheckBoxes = true;
		this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnFileName,
            this.columnFileDescription,
            this.columnHersteller,
            this.columnArtikelnummer});
		this.listView.ContextMenuStrip = this.contextMenuListView;
		this.listView.FullRowSelect = true;
		this.listView.HideSelection = false;
		this.listView.Location = new System.Drawing.Point(12, 12);
		this.listView.Name = "listView";
		this.listView.ShowGroups = false;
		this.listView.Size = new System.Drawing.Size(700, 300);
		this.listView.Sorting = System.Windows.Forms.SortOrder.Ascending;
		this.listView.TabIndex = 1;
		this.listView.UseCompatibleStateImageBehavior = false;
		this.listView.View = System.Windows.Forms.View.Details;
		this.listView.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listView_ColumnClick);
		// 
		// columnFileName
		// 
		this.columnFileName.Text = "Dokument";
		this.columnFileName.Width = 540;
		// 
		// columnFileDescription
		// 
		this.columnFileDescription.Text = "Beschreibung";
		this.columnFileDescription.Width = 150;
		// 
		// columnHersteller
		// 
		this.columnHersteller.Text = "Hersteller";
		this.columnHersteller.Width = 150;
		// 
		// columnArtikelnummer
		// 
		this.columnArtikelnummer.Text = "Artikelnummer";
		this.columnArtikelnummer.Width = 150;
		// 
		// contextMenuListView
		// 
		this.contextMenuListView.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemDocumentOpen});
		this.contextMenuListView.Name = "contextMenuListView";
		this.contextMenuListView.Size = new System.Drawing.Size(169, 26);
		this.contextMenuListView.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuListView_Opening);
		// 
		// toolStripMenuItemDocumentOpen
		// 
		this.toolStripMenuItemDocumentOpen.Name = "toolStripMenuItemDocumentOpen";
		this.toolStripMenuItemDocumentOpen.Size = new System.Drawing.Size(168, 22);
		this.toolStripMenuItemDocumentOpen.Text = "Dokument öffnen";
		this.toolStripMenuItemDocumentOpen.Click += new System.EventHandler(this.toolStripMenuItemDocumentOpen_Click);
		// 
		// btnKopieren
		// 
		this.btnKopieren.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnKopieren.Location = new System.Drawing.Point(457, 393);
		this.btnKopieren.Name = "btnKopieren";
		this.btnKopieren.Size = new System.Drawing.Size(120, 26);
		this.btnKopieren.TabIndex = 5;
		this.btnKopieren.Text = "&Kopieren";
		this.btnKopieren.UseVisualStyleBackColor = true;
		this.btnKopieren.Click += new System.EventHandler(this.btnKopieren_Click);
		// 
		// txtZielverzeichnis
		// 
		this.txtZielverzeichnis.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
		| System.Windows.Forms.AnchorStyles.Right)));
		this.txtZielverzeichnis.ContextMenuStrip = this.contextMenuZielverzeichnis;
		this.txtZielverzeichnis.Location = new System.Drawing.Point(12, 341);
		this.txtZielverzeichnis.Name = "txtZielverzeichnis";
		this.txtZielverzeichnis.Size = new System.Drawing.Size(669, 20);
		this.txtZielverzeichnis.TabIndex = 2;
		// 
		// contextMenuZielverzeichnis
		// 
		this.contextMenuZielverzeichnis.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuHerstellerVerzeichnis,
            this.toolStripMenuArtikelnummerVerzeichnis,
            this.toolStripSeparator1,
            this.toolStripMenuItemProjekteVerzeichnis,
            this.toolStripMenuItemProjektVerzeichnisstruktur});
		this.contextMenuZielverzeichnis.Name = "contextMenuZielverzeichnis";
		this.contextMenuZielverzeichnis.Size = new System.Drawing.Size(340, 98);
		// 
		// toolStripMenuHerstellerVerzeichnis
		// 
		this.toolStripMenuHerstellerVerzeichnis.CheckOnClick = true;
		this.toolStripMenuHerstellerVerzeichnis.Name = "toolStripMenuHerstellerVerzeichnis";
		this.toolStripMenuHerstellerVerzeichnis.Size = new System.Drawing.Size(339, 22);
		this.toolStripMenuHerstellerVerzeichnis.Text = "Hersteller zur Verzeichnisstruktur hinzufügen";
		// 
		// toolStripMenuArtikelnummerVerzeichnis
		// 
		this.toolStripMenuArtikelnummerVerzeichnis.CheckOnClick = true;
		this.toolStripMenuArtikelnummerVerzeichnis.Name = "toolStripMenuArtikelnummerVerzeichnis";
		this.toolStripMenuArtikelnummerVerzeichnis.Size = new System.Drawing.Size(339, 22);
		this.toolStripMenuArtikelnummerVerzeichnis.Text = "Artikelnummer zur Verzeichnisstruktur hinzufügen";
		// 
		// toolStripSeparator1
		// 
		this.toolStripSeparator1.Name = "toolStripSeparator1";
		this.toolStripSeparator1.Size = new System.Drawing.Size(336, 6);
		// 
		// toolStripMenuItemProjekteVerzeichnis
		// 
		this.toolStripMenuItemProjekteVerzeichnis.Name = "toolStripMenuItemProjekteVerzeichnis";
		this.toolStripMenuItemProjekteVerzeichnis.Size = new System.Drawing.Size(339, 22);
		this.toolStripMenuItemProjekteVerzeichnis.Text = "Projekt-Verzeichnisstruktur hinzufügen";
		this.toolStripMenuItemProjekteVerzeichnis.Click += new System.EventHandler(this.toolStripMenuItemProjekteVerzeichnis_Click);
		// 
		// toolStripMenuItemProjektVerzeichnisstruktur
		// 
		this.toolStripMenuItemProjektVerzeichnisstruktur.Name = "toolStripMenuItemProjektVerzeichnisstruktur";
		this.toolStripMenuItemProjektVerzeichnisstruktur.Size = new System.Drawing.Size(339, 22);
		this.toolStripMenuItemProjektVerzeichnisstruktur.Text = "Komplette Projekt-Verzeichnisstruktur hinzufügen";
		this.toolStripMenuItemProjektVerzeichnisstruktur.Click += new System.EventHandler(this.toolStripMenuItemProjektVerzeichnisstruktur_Click);
		// 
		// label1
		// 
		this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(9, 325);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(80, 13);
		this.label1.TabIndex = 4;
		this.label1.Text = "Zielverzeichnis:";
		// 
		// btnOrdnerWählen
		// 
		this.btnOrdnerWählen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnOrdnerWählen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
		this.btnOrdnerWählen.Location = new System.Drawing.Point(684, 338);
		this.btnOrdnerWählen.Name = "btnOrdnerWählen";
		this.btnOrdnerWählen.Size = new System.Drawing.Size(28, 24);
		this.btnOrdnerWählen.TabIndex = 3;
		this.btnOrdnerWählen.Text = "...";
		this.btnOrdnerWählen.UseVisualStyleBackColor = true;
		this.btnOrdnerWählen.Click += new System.EventHandler(this.btnOrdnerWählen_Click);
		// 
		// btnOdnerÖffnen
		// 
		this.btnOdnerÖffnen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnOdnerÖffnen.Location = new System.Drawing.Point(12, 393);
		this.btnOdnerÖffnen.Name = "btnOdnerÖffnen";
		this.btnOdnerÖffnen.Size = new System.Drawing.Size(120, 26);
		this.btnOdnerÖffnen.TabIndex = 7;
		this.btnOdnerÖffnen.Text = "Verzeichnis &öffnen";
		this.btnOdnerÖffnen.UseVisualStyleBackColor = true;
		this.btnOdnerÖffnen.Click += new System.EventHandler(this.btnOdnerÖffnen_Click);
		// 
		// btnExtras
		// 
		this.btnExtras.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnExtras.Location = new System.Drawing.Point(318, 393);
		this.btnExtras.Name = "btnExtras";
		this.btnExtras.Size = new System.Drawing.Size(120, 26);
		this.btnExtras.TabIndex = 4;
		this.btnExtras.Text = "E&xtras";
		this.btnExtras.UseVisualStyleBackColor = true;
		this.btnExtras.Click += new System.EventHandler(this.btnExtras_Click);
		// 
		// frmDocumentationTool
		// 
		this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.CancelButton = this.btnAbbrechen;
		this.ClientSize = new System.Drawing.Size(726, 431);
		this.Controls.Add(this.btnExtras);
		this.Controls.Add(this.btnOdnerÖffnen);
		this.Controls.Add(this.btnOrdnerWählen);
		this.Controls.Add(this.label1);
		this.Controls.Add(this.txtZielverzeichnis);
		this.Controls.Add(this.btnKopieren);
		this.Controls.Add(this.listView);
		this.Controls.Add(this.btnAbbrechen);
		this.MinimizeBox = false;
		this.Name = "frmDocumentationTool";
		this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
		this.Text = "Documentation-Tool";
		this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmDocumentationTool_FormClosing);
		this.Load += new System.EventHandler(this.frmDocumentationTool_Load);
		this.contextMenuListView.ResumeLayout(false);
		this.contextMenuZielverzeichnis.ResumeLayout(false);
		this.ResumeLayout(false);
		this.PerformLayout();

	}

	public frmDocumentationTool()
	{
		InitializeComponent();
	}

	#endregion

	private int sortColumn = -1;

	//Skript wird entladen
	[DeclareUnregister]
	public void OnUnRegisterScript()
	{
		SettingsDelete();
	}

	//[DeclareMenu()]
	//public void Menupunkt()
	//{
	//	Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
	//	oMenu.AddMenuItem("Dokumentations-Tool...", "Documentation_Tool_Start", "Externe Dokumente ermitteln und kopieren", 35379, 1, false, false);
	//}

	[Start]
	//[DeclareAction("Documentation_Tool_Start")]
	public void Function()
	{
		frmDocumentationTool frm = new frmDocumentationTool();
		frm.ShowDialog();

		return;
	}

	//Form_Load
	private void frmDocumentationTool_Load(object sender, System.EventArgs e)
	{
		//Position und Größe aus Settings lesen
#if !DEBUG
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top"))
		{
			this.Top = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left"))
		{
			this.Left = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height"))
		{
			this.Height = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width"))
		{
			this.Width = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuHerstellerVerzeichnis"))
		{
			this.toolStripMenuHerstellerVerzeichnis.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuHerstellerVerzeichnis", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuArtikelnummerVerzeichnis"))
		{
			this.toolStripMenuArtikelnummerVerzeichnis.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuArtikelnummerVerzeichnis", 0);
		}
#endif

		//Titelzeile anpassen
		string sProjekt = string.Empty;
#if DEBUG
		sProjekt = "DEBUG";
#else
		CommandLineInterpreter cmdLineItp = new CommandLineInterpreter();
		ActionCallingContext ProjektContext = new ActionCallingContext();
		ProjektContext.AddParameter("TYPE", "PROJECT");
		cmdLineItp.Execute("selectionset", ProjektContext);
		ProjektContext.GetParameter("PROJECT", ref sProjekt);
		sProjekt = Path.GetFileNameWithoutExtension(sProjekt); //Projektname Pfad und ohne .elk
		if (sProjekt == string.Empty)
		{
			Decider eDecision = new Decider();
			EnumDecisionReturn eAnswer = eDecision.Decide(EnumDecisionType.eOkDecision, "Es ist kein Projekt ausgewählt.", "Documentation-Tool", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
			if (eAnswer == EnumDecisionReturn.eOK)
			{
				Close();
				return;
			}
		}
#endif
		Text = Text + " - " + sProjekt;

		//Button Extras Text festlegen
		//btnExtras.Text = "            Extras           ▾"; // ▾ ▼

		//Zielverzeichnis vorbelegen
#if DEBUG
		txtZielverzeichnis.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Test";
#else
		txtZielverzeichnis.Text = PathMap.SubstitutePath(@"$(DOC)");
#endif

		// Temporären Dateinamen festlegen
#if DEBUG
		string sTempFile = Path.Combine(Application.StartupPath, "tmp_Projekt_Export.epj");
#else
		string sTempFile = Path.Combine(PathMap.SubstitutePath(@"$(TMP)"), "tmpDocumentationTool.epj");
#endif

		//Projekt exportieren
#if !DEBUG
		PXFexport(sTempFile);
#endif

		//PXF Datei einlesen und in Listview schreiben
		PXFeinlesen(sTempFile);

		//PXF Datei wieder löschen
#if !DEBUG
		File.Delete(sTempFile);
#endif
	}

	//Form_Close 
	private void frmDocumentationTool_FormClosing(object sender, FormClosingEventArgs e)
	{
		//Einstellungen speichern
		SettingsWrite();
	}

	//Gesamtes Projekt als PXF ausgeben
	public void PXFexport(string sFilename)
	{
		//Progressbar ein
		Progress oProgress = new Progress("SimpleProgress");
		oProgress.SetTitle("Documentation-Tool");
		oProgress.SetActionText("Projektdaten exportieren");
		oProgress.ShowImmediately();
		Application.DoEvents(); // Screen aktualisieren

		ActionCallingContext pxfContext = new ActionCallingContext();
		pxfContext.AddParameter("TYPE", "PXFPROJECT");
		pxfContext.AddParameter("EXPORTFILE", sFilename);
		pxfContext.AddParameter("EXPORTMASTERDATA", "0"); //Stammdaten mit exportieren (Standard = 1(Ja)
		pxfContext.AddParameter("EXPORTCONNECTIONS", "0"); //Verbindungen mit exportieren (Standard = 0(Nein)

		CommandLineInterpreter cmdLineItp = new CommandLineInterpreter();
		cmdLineItp.Execute("export", pxfContext);

		//Progressbar aus
		oProgress.EndPart(true);

		return;
	}

	//PXF Datei einlesen
	private void PXFeinlesen(string sFileName)
	{
		//Progressbar ein
#if !DEBUG
		Progress oProgress = new Progress("SimpleProgress");
		oProgress.SetTitle("Documentation-Tool");
		//oProgress.SetActionText("Projektdaten durchsuchen");
		oProgress.BeginPart(50,"Projektdaten durchsuchen");
		oProgress.ShowImmediately();
#endif

		//MessageBox.Show("XML.Reader :" + sFileName);
		ListViewItem objListViewItem = new ListViewItem();

		//Wir benötigen einen XmlReader für das Auslesen der XML-Datei 
		XmlTextReader XMLReader = new XmlTextReader(sFileName);

		//Es folgt das Auslesen der XML-Datei 
		while (XMLReader.Read()) //Es sind noch Daten vorhanden 
		{
			//Alle Attribute (Name-Wert-Paare) abarbeiten 
			if (XMLReader.AttributeCount > 0)
			{
				string sArtikelnummer = string.Empty;
				string sHersteller = string.Empty;
				//Es sind noch weitere Attribute vorhanden 
				while (XMLReader.MoveToNextAttribute()) //nächstes
				{
					if (XMLReader.Name == "P22001") // Artikel (eingelagert) Artikelnummer)
					{
						sArtikelnummer = XMLReader.Value;
					}
					if (XMLReader.Name == "P22007") // Artikel (eingelagert) Hersteller)
					{
						sHersteller = XMLReader.Value;
					}
					if (
						XMLReader.Name == "A2082" || // Hyperlink Dokument
						//XMLReader.Name == "P11058" || // Fremddokument
						XMLReader.Name == "P22149" || // Artikel (eingelagert) Externes Dokument 1
						XMLReader.Name == "P22150" || // Artikel (eingelagert) Externes Dokument 2
						XMLReader.Name == "P22151" || // Artikel (eingelagert) Externes Dokument 3
						XMLReader.Name == "P22235" || // Artikel (eingelagert) Externes Dokument 4
						XMLReader.Name == "P22236" || // Artikel (eingelagert) Externes Dokument 5
						XMLReader.Name == "P22237" || // Artikel (eingelagert) Externes Dokument 6
						XMLReader.Name == "P22238" || // Artikel (eingelagert) Externes Dokument 7
						XMLReader.Name == "P22239" || // Artikel (eingelagert) Externes Dokument 8
						XMLReader.Name == "P22240" || // Artikel (eingelagert) Externes Dokument 9
						XMLReader.Name == "P22241" || // Artikel (eingelagert) Externes Dokument 10
						XMLReader.Name == "P22242" || // Artikel (eingelagert) Externes Dokument 11
						XMLReader.Name == "P22243" || // Artikel (eingelagert) Externes Dokument 12
						XMLReader.Name == "P22244" || // Artikel (eingelagert) Externes Dokument 13
						XMLReader.Name == "P22245" || // Artikel (eingelagert) Externes Dokument 14
						XMLReader.Name == "P22246" || // Artikel (eingelagert) Externes Dokument 15
						XMLReader.Name == "P22247" || // Artikel (eingelagert) Externes Dokument 16
						XMLReader.Name == "P22248" || // Artikel (eingelagert) Externes Dokument 17
						XMLReader.Name == "P22249" || // Artikel (eingelagert) Externes Dokument 18
						XMLReader.Name == "P22250" || // Artikel (eingelagert) Externes Dokument 19
						XMLReader.Name == "P22251" // Artikel (eingelagert) Externes Dokument 20
						)
					{

						string[] sValue = XMLReader.Value.Split('\n');

						string sDateiname = string.Empty;
						string sDateiBeschreibung = string.Empty;
						sDateiname = sValue[0];

						//Überprüfen ob Beschreibung vorhanden ist
						if (sValue.Length == 2)
						{
							if (sValue[1] != null && sValue[1] != string.Empty)
							{
								sDateiBeschreibung = sValue[1];
#if !DEBUG
								MultiLangString mlstrDateiBeschreibung = new MultiLangString();
								//Nur die deutsche Übersetzung verwenden
								//String in MultiLanguages wandeln
								mlstrDateiBeschreibung.SetAsString(sDateiBeschreibung);
								//Daraus nur die Deutsche übersetzung
								sDateiBeschreibung = mlstrDateiBeschreibung.GetString(ISOCode.Language.L_de_DE);
								//Wenn es keine Deutsche gibt, dann die unbestimmte
								if (sDateiBeschreibung == "")
								{
									sDateiBeschreibung = mlstrDateiBeschreibung.GetString(ISOCode.Language.L___);
								}
#endif
							}
						}

#if !DEBUG
						sDateiname = PathMap.SubstitutePath(sDateiname);
#endif
						//keine Dokumente die mit HTTP anfangen bearbeiten
						if (!sDateiname.StartsWith("http"))
						{
							objListViewItem = new ListViewItem();
							objListViewItem.Name = sDateiname; // Name muß gesetzt werden damit ContainsKey funktioniert
							objListViewItem.Text = sDateiname;
							objListViewItem.SubItems.Add(sDateiBeschreibung); //Datei Beschreibung
							objListViewItem.SubItems.Add(sHersteller); //Hersteller
							objListViewItem.SubItems.Add(sArtikelnummer); //Artikelnummer
							objListViewItem.Checked = true;

							//Prüfen ob Datei vorhanden
							if (!File.Exists(sDateiname))
							{
								objListViewItem.Checked = false;
								objListViewItem.ForeColor = System.Drawing.Color.Gray;
							}

							//Eintrag in Listview, wenn nicht schon vorhanden
							if (!listView.Items.ContainsKey(sDateiname))
							{
								listView.Items.Add(objListViewItem);
							}
						}
					}
				}
			}
		}

		//XMLTextReader schließen
		XMLReader.Close();

		//Spaltenbreite automatisch an Inhaltsbreite anpassen
		//listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);

		//Progressbar aus
#if !DEBUG
		oProgress.EndPart(true);
#endif

		return;
	}

	//Dokument kopieren
	public void DocumentCopy(string sDokument, string sZiel)
	{
		//gibt es das Dokument auch?
		if (File.Exists(sDokument))
		{
			//Dateinamen ermitteln
			string sDateiname = Path.GetFileName(sDokument);

			try
			{
				//Dokument kopieren, mit überschreiben
				File.Copy(sDokument, Path.Combine(sZiel, sDateiname), true);
			}
			catch (Exception exc)
			{
				String strMessage = exc.Message;
				MessageBox.Show("Exception: " + strMessage + "\n\n" + sZiel + " -- " + sDateiname, "Documentation-Tool, DocumentCopy", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		else
		{
			//Hinweis, Datei gibt es nicht
#if DEBUG
			MessageBox.Show("Das Dokument [" + sDokument + "] ist nicht auf dem Datenträger vorhanden!", "Documentation-Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
#else
			Decider eDecision = new Decider();
			eDecision.Decide(EnumDecisionType.eOkDecision, "Das Dokument [" + sDokument + "] ist nicht auf dem Datenträger vorhanden!", "Documentation-Tool", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);

#endif
		}
		return;
	}

	//Ordner auswählen
	public string OrdnerAuswählen(string InitialDirectory)
	{
		var folderBrowser = new FolderBrowserDialog();
		folderBrowser.Description = "Bitte wählen Sie einen Ordner aus";
		folderBrowser.SelectedPath = InitialDirectory;
		folderBrowser.ShowNewFolderButton = true;
		DialogResult result = folderBrowser.ShowDialog();
		if (result == DialogResult.OK)
		{
			return folderBrowser.SelectedPath;
		}
		else
		{
			return string.Empty;
		}
	}

	//Button: Abbrechen
	private void btnAbbrechen_Click(object sender, System.EventArgs e)
	{
		Close();
	}

	//Button: Zielverzeichniss auswählen
	private void btnOrdnerWählen_Click(object sender, EventArgs e)
	{
		string sZielverzeichnis = txtZielverzeichnis.Text;
		sZielverzeichnis = OrdnerAuswählen(txtZielverzeichnis.Text);
		if (sZielverzeichnis != string.Empty)
		{
			txtZielverzeichnis.Text = sZielverzeichnis;
		}
	}

	//Button: Zielverzeichnis im Explorer öffnen
	private void btnOdnerÖffnen_Click(object sender, EventArgs e)
	{
		//Start Windows-Explorer mit Parameter
		System.Diagnostics.Process.Start("explorer", "/e," + txtZielverzeichnis.Text);
		return;
	}

	//Button: Dokumente kopieren
	private void btnKopieren_Click(object sender, EventArgs e)
	{
		//gibt es das Ziel auch?
		if (!Directory.Exists(txtZielverzeichnis.Text))
		{
			//Hinweis, Ziel gibt es nicht, anlegen?
#if DEBUG
			DialogResult dialogResult = MessageBox.Show("Das Zielverzeichnis [" + txtZielverzeichnis.Text + "] ist nicht auf dem Datenträger vorhanden!\n\nSoll dieses nun angelegt werden?", "Documentation-Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
			if (dialogResult == DialogResult.Yes)
			{
				Directory.CreateDirectory(txtZielverzeichnis.Text);
			}
			else if (dialogResult == DialogResult.No)
			{
				return;
			}
#else
			Decider decider = new Decider();
			EnumDecisionReturn decision = decider.Decide(
				EnumDecisionType.eYesNoDecision, // type
				"Das Zielverzeichnis [" + txtZielverzeichnis.Text + "] ist nicht auf dem Datenträger vorhanden!\n\nSoll dieses nun angelegt werden?",
				"Documentation-Tool",
				EnumDecisionReturn.eYES, // selected Answer
				EnumDecisionReturn.eYES); // Answer if quite-mode on

			if (decision == EnumDecisionReturn.eYES)
			{
				Directory.CreateDirectory(txtZielverzeichnis.Text);
			}
			else if (decision == EnumDecisionReturn.eNO)
			{
				return;
			}
#endif
		}

		//Es gibt dasZiel
		if (Directory.Exists(txtZielverzeichnis.Text))
		{
			ListView.CheckedListViewItemCollection checkedItems = listView.CheckedItems;
#if !DEBUG
		Progress oProgress = new Progress("SimpleProgress");
		oProgress.ShowImmediately();
		oProgress.SetAllowCancel(true);
		oProgress.SetTitle("Documentation-Tool");
		int nActionsPercent = 100 / checkedItems.Count;
#endif
			foreach (ListViewItem item in checkedItems)
			{
#if !DEBUG
			if (!oProgress.Canceled())
			{
				oProgress.BeginPart(nActionsPercent, "Kopiere: " + item.Text);
#endif
				string sTargetPath = txtZielverzeichnis.Text;

				//Hersteller-Verzeichnis hinzufügen
				if (toolStripMenuHerstellerVerzeichnis.Checked)
				{
					try
					{
						string sTemp = item.SubItems[2].Text;
						sTemp = RemoveIlegaleCharackter(sTemp); //Verzeichnisnamen von ungültigen Zeichen bereinigen
						sTargetPath = Path.Combine(sTargetPath, sTemp);
						if (!Directory.Exists(sTargetPath)) //Verzeichnis anlegen wenn noch nicht vorhanden
						{
							Directory.CreateDirectory(sTargetPath);
						}
					}
					catch (Exception exc)
					{
						String strMessage = exc.Message;
						MessageBox.Show("Exception: " + strMessage + "\n\n" + sTargetPath, "Documentation-Tool, btnKopieren", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}

				//Artikelnummer-Verzeichnis hinzufügen
				if (toolStripMenuArtikelnummerVerzeichnis.Checked)
				{
					try
					{
						string sTemp = item.SubItems[3].Text;
						sTemp = RemoveIlegaleCharackter(sTemp); //Verzeichnisnamen von ungültigen Zeichen bereinigen
						sTargetPath = Path.Combine(sTargetPath, sTemp);

						if (!Directory.Exists(sTargetPath)) //Verzeichnis anlegen wenn noch nicht vorhanden
						{
							Directory.CreateDirectory(sTargetPath);
						}
					}
					catch (Exception exc)
					{
						String strMessage = exc.Message;
						MessageBox.Show("Exception: " + strMessage + "\n\n" + sTargetPath, "Documentation-Tool, btnKopieren", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}

				DocumentCopy(item.Text, sTargetPath);

#if !DEBUG
			oProgress.EndPart();
			}
#endif
			}
#if !DEBUG
		oProgress.EndPart(true);
#endif
			Close();
			return;
		}
	}

	//Button: Extras (Kontextmenü Zielverzeichnis anzeigen)
	private void btnExtras_Click(object sender, EventArgs e)
	{
		contextMenuZielverzeichnis.Show(btnExtras, 0, 0 - (contextMenuZielverzeichnis.Height - btnExtras.Height));
	}

	//Listview: Spalten angeklickt
	private void listView_ColumnClick(object sender, ColumnClickEventArgs e)
	{
		// Determine whether the column is the same as the last column clicked.
		if (e.Column != sortColumn)
		{
			// Set the sort column to the new column.
			sortColumn = e.Column;
			// Set the sort order to ascending by default.
			listView.Sorting = SortOrder.Ascending;
		}
		else
		{
			// Determine what the last sort order was and change it.
			if (listView.Sorting == SortOrder.Ascending)
				listView.Sorting = SortOrder.Descending;
			else
				listView.Sorting = SortOrder.Ascending;
		}

		// Call the sort method to manually sort.
		listView.Sort();
		// Set the ListViewItemSorter property to a new ListViewItemComparer
		// object.
		this.listView.ListViewItemSorter = new ListViewItemComparer(e.Column, listView.Sorting);
	}

	//Kontextmenü 'Projekt-Verzeichnisstruktur hinzufügen'
	private void toolStripMenuItemProjekteVerzeichnis_Click(object sender, EventArgs e)
	{
		string sProjectePath = string.Empty;
		string sSelectedProjectPath = string.Empty;
#if DEBUG
		sProjectePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
		sSelectedProjectPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Test\Test.edb";
#else
		sProjectePath = PathMap.SubstitutePath(@"$(MD_PROJECTS)");
		sSelectedProjectPath = PathMap.SubstitutePath(@"$(P)");
#endif
		if (sProjectePath != string.Empty && sSelectedProjectPath != string.Empty)
		{
			string sTemp = string.Empty;

			sTemp = sSelectedProjectPath.Replace(sProjectePath, string.Empty); //Projekte-Verzeichnis aus Projekt-Verzeichnis entfernen
			sTemp = Path.ChangeExtension(sTemp, null); //Erweiterung entfernen

			//Überprüfen ob mit "\" beginnt, dann entfernen			
			if (Path.IsPathRooted(sTemp))
			{
				sTemp = sTemp.TrimStart(Path.DirectorySeparatorChar);
				sTemp = sTemp.TrimStart(Path.AltDirectorySeparatorChar);
			}

			if (!txtZielverzeichnis.Text.EndsWith(@"\"))
			{
				txtZielverzeichnis.Text += @"\";
			}

			//neues Zielverzeichnis eintragen
			txtZielverzeichnis.Text = Path.Combine(txtZielverzeichnis.Text, sTemp);
			txtZielverzeichnis.Select(); // to Set Focus
			txtZielverzeichnis.Select(txtZielverzeichnis.Text.Length, 0); //to set cursor at the end of textbox
		}
	}

	//Kontextmenü 'Komplette Projekt-Verzeichnisstruktur hinzufügen'
	private void toolStripMenuItemProjektVerzeichnisstruktur_Click(object sender, EventArgs e)
	{
		string sSelectedProjectPath = string.Empty;
#if DEBUG
		sSelectedProjectPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Test\Test.edb";
#else
		sSelectedProjectPath = PathMap.SubstitutePath(@"$(P)");
#endif
		if (sSelectedProjectPath != string.Empty)
		{
			string sTemp = string.Empty;

			string sRootPath = Path.GetPathRoot(sSelectedProjectPath); //Root-Verzeichnis ermitteln
			sTemp = sSelectedProjectPath.Replace(sRootPath, string.Empty); //Root-Verzeichnis aus Projekt-Verzeichnis entfernen
			sTemp = Path.ChangeExtension(sTemp, null); //Erweiterung entfernen

			//Überprüfen ob mit "\" beginnt, dann entfernen			
			if (Path.IsPathRooted(sTemp))
			{
				sTemp = sTemp.TrimStart(Path.DirectorySeparatorChar);
				sTemp = sTemp.TrimStart(Path.AltDirectorySeparatorChar);
			}

			if (!txtZielverzeichnis.Text.EndsWith(@"\"))
			{
				txtZielverzeichnis.Text += @"\";
			}

			//neues Zielverzeichnis eintragen
			txtZielverzeichnis.Text = Path.Combine(txtZielverzeichnis.Text, sTemp);
			txtZielverzeichnis.Select(); // to Set Focus
			txtZielverzeichnis.Select(txtZielverzeichnis.Text.Length, 0); //to set cursor at the end of textbox
		}
	}

	//Kontextmenü listView öffnen
	private void contextMenuListView_Opening(object sender, System.ComponentModel.CancelEventArgs e)
	{
		if (listView.SelectedItems.Count != 0)
		{
			if (listView.SelectedItems[0].ForeColor == System.Drawing.Color.Gray)
			{
				toolStripMenuItemDocumentOpen.Enabled = false;
			}
			else
			{
				toolStripMenuItemDocumentOpen.Enabled = true;
			}
		}
		else
		{
			toolStripMenuItemDocumentOpen.Enabled = false;
		}
		return;
	}

	//Kontextmenü 'Dokument öffnen'
	private void toolStripMenuItemDocumentOpen_Click(object sender, EventArgs e)
	{
		if (listView.SelectedItems.Count == 1)
		{
			System.Diagnostics.Process.Start(listView.SelectedItems[0].Text);
			return;
		}
	}

	//Einstellungen speichen
	private void SettingsWrite()
	{
#if !DEBUG
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top",
			new int[] { },
			new Range[] { },
			ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top", this.Top, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left",
				new int[] { },
				new Range[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left", this.Left, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height",
				new int[] { },
				new Range[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height", this.Height, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width",
				new int[] { },
				new Range[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width", this.Width, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuHerstellerVerzeichnis"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuHerstellerVerzeichnis",
				new bool[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuHerstellerVerzeichnis", toolStripMenuHerstellerVerzeichnis.Checked, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuArtikelnummerVerzeichnis"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuArtikelnummerVerzeichnis",
				new bool[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.toolStripMenuArtikelnummerVerzeichnis", toolStripMenuArtikelnummerVerzeichnis.Checked, 0);
#endif
		return;
	}

	//Einstellungen löschen
	public void SettingsDelete()
	{
		DialogResult Result = MessageBox.Show("Sollen die Einstellungen wirklich aus EPLAN gelöscht werden?", "Documentation-Tool, Einstellungen löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
		if (Result == System.Windows.Forms.DialogResult.Yes)
		{
			//Settings löschen
			Eplan.EplApi.Base.SettingNode oSettingNode = new Eplan.EplApi.Base.SettingNode("USER.SCRIPTS.DOCUMENTATION_TOOL");

			MessageBox.Show("Es wurden " + oSettingNode.GetCountOfSettings().ToString() + " Einstellungen gelöscht.", "Documentation-Tool, Einstellungen löschen", MessageBoxButtons.OK, MessageBoxIcon.Information);
			oSettingNode.DeleteNode();
		}
		return;
	}

	//ungültige Zeichen aus Dateinamen entfernen
	private static string RemoveIlegaleCharackter(string fileName)
	{
		string illegal = fileName;
		string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
		
		foreach (char c in invalid)
		{
			illegal = illegal.Replace(c.ToString(), "");
		}

		return illegal;
	}
}

// Implements the manual sorting of items by columns.
class ListViewItemComparer : IComparer
{
	private int col;
	private SortOrder order;
	public ListViewItemComparer()
	{
		col = 0;
		order = SortOrder.Ascending;
	}
	public ListViewItemComparer(int column, SortOrder order)
	{
		col = column;
		this.order = order;
	}
	public int Compare(object x, object y)
	{
		int returnVal = -1;
		returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
								((ListViewItem)y).SubItems[col].Text);
		// Determine whether the sort order is descending.
		if (order == SortOrder.Descending)
			// Invert the value returned by String.Compare.
			returnVal *= -1;
		return returnVal;
	}
}