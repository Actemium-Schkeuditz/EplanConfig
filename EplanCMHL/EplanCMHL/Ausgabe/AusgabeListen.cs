using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;


namespace EplanCMHL.Ausgabe
{

    public partial class AusgabeListen : Form
    {

        private System.Windows.Forms.Button ButtonSummaryPartlist;
        private System.Windows.Forms.Button ButtonCancel;
        private System.Windows.Forms.Label label1;
        private Label label2;
        private GroupBox groupBoxLocations;
        private CheckBox checkBoxSummery;
        private Button button1;
        private Button button2;

        #region Windows Form Designer generated code
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }



        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ButtonSummaryPartlist = new System.Windows.Forms.Button();
            this.ButtonCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBoxLocations = new System.Windows.Forms.GroupBox();
            this.checkBoxSummery = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBoxLocations.SuspendLayout();
            this.SuspendLayout();
            // 
            // ButtonSummaryPartlist
            // 
            this.ButtonSummaryPartlist.Location = new System.Drawing.Point(469, 12);
            this.ButtonSummaryPartlist.Name = "ButtonSummaryPartlist";
            this.ButtonSummaryPartlist.Size = new System.Drawing.Size(116, 23);
            this.ButtonSummaryPartlist.TabIndex = 10;
            this.ButtonSummaryPartlist.Text = "Summenstückliste";
            this.ButtonSummaryPartlist.UseVisualStyleBackColor = true;
            this.ButtonSummaryPartlist.Click += new System.EventHandler(this.ButtonSummaryPartlist_Click);
            // 
            // ButtonCancel
            // 
            this.ButtonCancel.Location = new System.Drawing.Point(1002, 12);
            this.ButtonCancel.Name = "ButtonCancel";
            this.ButtonCancel.Size = new System.Drawing.Size(75, 23);
            this.ButtonCancel.TabIndex = 11;
            this.ButtonCancel.Text = "Beenden";
            this.ButtonCancel.UseVisualStyleBackColor = true;
            this.ButtonCancel.Click += new System.EventHandler(this.ButtonCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(550, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Einbauorte";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Anlagen";
            // 
            // groupBoxLocations
            // 
            this.groupBoxLocations.AutoSize = true;
            this.groupBoxLocations.Controls.Add(this.label1);
            this.groupBoxLocations.Controls.Add(this.label2);
            this.groupBoxLocations.Location = new System.Drawing.Point(12, 46);
            this.groupBoxLocations.Name = "groupBoxLocations";
            this.groupBoxLocations.Size = new System.Drawing.Size(1097, 200);
            this.groupBoxLocations.TabIndex = 14;
            this.groupBoxLocations.TabStop = false;
            this.groupBoxLocations.Text = "Auswahl der Anlagen und Einbauorte";
            // 
            // checkBoxSummery
            // 
            this.checkBoxSummery.AutoSize = true;
            this.checkBoxSummery.Location = new System.Drawing.Point(791, 16);
            this.checkBoxSummery.Name = "checkBoxSummery";
            this.checkBoxSummery.Size = new System.Drawing.Size(121, 17);
            this.checkBoxSummery.TabIndex = 14;
            this.checkBoxSummery.Text = "Auswahl Summieren";
            this.checkBoxSummery.UseVisualStyleBackColor = true;
            this.checkBoxSummery.CheckedChanged += new System.EventHandler(this.Box_SummaryChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(330, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(116, 23);
            this.button1.TabIndex = 15;
            this.button1.Text = "Artikelstückliste";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.ButtonPartlist_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(197, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(116, 23);
            this.button2.TabIndex = 16;
            this.button2.Text = "Kabellisten";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.ButtonCablelist_Click);
            // 
            // AusgabeListen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1128, 564);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkBoxSummery);
            this.Controls.Add(this.groupBoxLocations);
            this.Controls.Add(this.ButtonSummaryPartlist);
            this.Controls.Add(this.ButtonCancel);
            this.MinimumSize = new System.Drawing.Size(500, 400);
            this.Name = "AusgabeListen";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Ausgabe von Listen";
            this.Load += new System.EventHandler(this.LoadFrmSelect);
            this.groupBoxLocations.ResumeLayout(false);
            this.groupBoxLocations.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public AusgabeListen()
            {
            InitializeComponent();
        }

        #endregion

        private List<string> sAnlagenWahl = new List<String>();
        private List<string> sEinbauortWahl = new List<String>();
        private string sTagEinbauort = "Einbauort";
        private string sTagAnlagen = "Anlagen";
        private bool getSummaryList = false;

        string strLabelName = PathMap.SubstitutePath("$(PROJECTNAME)");
        string myDate = System.DateTime.Now.ToString("yyyyMMdd");
        string myTime = System.DateTime.Now.ToString("HHmm");
        string sProjectDocPath = PathMap.SubstitutePath("$(DOC)");

       // [DeclareAction("AusgabeSummenstueckListe")]
        [Start]
        public void MenueAusgabeEinbauOrte()
        {

            // Projektnamen holen
            string sProject = string.Empty;
            ActionCallingContext oCTX1 = new ActionCallingContext();
            CommandLineInterpreter oCLI1 = new CommandLineInterpreter();
            oCTX1.AddParameter("TYPE", "PROJECT");
            oCLI1.Execute("selectionset", oCTX1);
            oCTX1.GetParameter("PROJECT", ref sProject);

            if (sProject != string.Empty)
            {
                // Formular aufrufen
                FormAuswahlDaten();
                //MessageBox.Show("Aktuelles Projekt: " + sProject, "Action: selectionset");
            }
            else
            {
                MessageBox.Show("Kein Projekt geöffnet !", "Action: selectionset");
            }

        }

        public void FormAuswahlDaten()
        {
            // Formular aufrufen
            AusgabeListen ausgabeAuswahlDaten = new AusgabeListen();
            AusgabeListen form = ausgabeAuswahlDaten;           
            form.ShowDialog();
            return;
        }

        private void CreateBoxString(List<string> daten, string sTag, int iY, int iMax)
        {
            int x = 0;
            int y = 0;
            CheckBox box;
            for (int i = 0; i < daten.Count(); i++)
            {
              // MessageBox.Show(daten[x].ToString()); //test
                box = new CheckBox();
                box.CheckedChanged += Box_CheckedChanged; // here
                box.Text = daten[i].ToString();
                box.Tag = sTag;
                box.AutoSize = true;
                x = i * 20;
                if (x > (2 * iMax))
                {
                    x = x - (2 * iMax) + 20;
                    y = 5 + iY + 340;
                }
                else if (x > iMax)
                {
                    x = x - iMax + 20;
                    y = 5 + iY + 190;
                }
                else   // erste Reihe
                {
                    y = 5 + iY;
                    x += 40;
                }
                //  x += 40; // Offset von oben
                box.Location = new Point(y, x);
                box.Checked = false;
                this.groupBoxLocations.Controls.Add(box);
            }
        }
        private void Box_CheckedChanged(object objSender, EventArgs e)
        {
            // 
            CheckBox cb = objSender as CheckBox; // get a reference to the checked/unchecked CheckBox

            // Abfragen ob gesetzt
            if (cb.Checked)
            {
                //MessageBox.Show(cb.Text + " checked");
                if (cb.Tag.ToString() == sTagEinbauort)
                {
                    //MessageBox.Show(cb.Text + " Einbauort checked");
                    sEinbauortWahl.Add(cb.Text);
                }
                else
                {
                  //  MessageBox.Show(cb.Text + " Anlage checked");
                    sAnlagenWahl.Add(cb.Text);
                }
            
            }
            else
            {
                if (cb.Tag.ToString() == sTagEinbauort)
                {
                    //MessageBox.Show(cb.Text + " Einbauort unchecked");
                    sEinbauortWahl.Remove(cb.Text);
                }
                else
                {
                   // MessageBox.Show(cb.Text + " Anlage unchecked");
                    sAnlagenWahl.Remove(cb.Text);
                }
                
            }
        }


        private void ButtonCancel_Click(object sender, EventArgs e)
        {

            this.Close();
        }

        private void ButtonSummaryPartlist_Click(object sender, EventArgs e)
        {
            string sSchemename = "ExportFunc_Artikelsummenstueckliste";
            string SETTINGS_PATH = "USER.Labelling.Config";
            string sImportFileArtikelstueckliste = @"$(MD_SCRIPTS)\Ausgabe\LB.ExportFunc_Artikelsummenstueckliste.xml";
            
            bool result = ImportScheme(sSchemename, SETTINGS_PATH, sImportFileArtikelstueckliste);

            if (!result)
            {
                MessageBox.Show("Fehler beim Import der Filter");
                return;
            }


            Progress progress = new Progress("EnhancedProgress");
            progress.SetAllowCancel(false);
            progress.ShowImmediately();
            progress.SetTitle("Ausgabe Stücklisten");

            try
            {

                string sSummaryLocation;
                string sSummaryHigherLevel;
                string sExportFilterSettings;
                string sExportFilename;
                string sSortscheme = "";
                // Prüfen ob eine Wahl erfolgt ist
                if ((sEinbauortWahl.Count == 0) & (sAnlagenWahl.Count == 0))
                {
                    MessageBox.Show("Bitte einen Einbauort oder eine Anlage wählen");
                    return;
                }

                // Ausgabe Summenstückliste
                if (getSummaryList)
                {
                    progress.BeginPart(33, "Summenliste");
                    // aus der Liste einen String machen
                    sSummaryHigherLevel = string.Join(";", sAnlagenWahl.ToArray());
                    sSummaryLocation = string.Join(";", sEinbauortWahl.ToArray());

                    if (!(string.IsNullOrEmpty(sSummaryHigherLevel)) & !(string.IsNullOrEmpty(sSummaryLocation)))
                    {
                        //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
                        //MessageBox.Show("Anlage und Einbauort" + sSummaryHigherLevel + sSummaryLocation);
                        sExportFilterSettings = "0|1|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Artikelsummenliste_" + strLabelName + "_" + sSummaryHigherLevel.Replace(";", "_") + "_" + sSummaryLocation.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei: 
                    }
                    else if (string.IsNullOrEmpty(sSummaryLocation))
                    {
                        //                          nur Anlage                                                                0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0| S1.F3 |0|1|1|0|0|0;0|
                        //MessageBox.Show("Anlage");
                        sExportFilterSettings = "0|1|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Artikelsummenliste_" + strLabelName + "_" + sSummaryHigherLevel.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:
                    }
                    else //(string.IsNullOrEmpty(sSummaryHigherLevel))
                    {
                        //                          nur Einbauort                                                             0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|   S1.F3          |0|1|1|0|0|0;0|
                        //MessageBox.Show("nur Einbauort");
                        sExportFilterSettings = "0|0|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Artikelsummenliste_" + strLabelName + "_" + sSummaryLocation.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:

                    }

                    ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                    progress.EndPart();
                }
                else
                {
                    // Progressbar zusammen bauen
                    double iSteps = 1;

                    if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count > 0))
                        iSteps = 1 / (sEinbauortWahl.Count * sAnlagenWahl.Count);
                    else if (sAnlagenWahl.Count > 0)
                    {
                        iSteps = 1 / sAnlagenWahl.Count;
                    }
                    else if (sEinbauortWahl.Count > 0)
                    {
                        iSteps = 1 / sEinbauortWahl.Count;
                    }                    

                    // Ausgabe der Daten in einzelne Listen

                    if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count > 0))
                    {
                        foreach (var ItemAnlage in sAnlagenWahl)
                        {
                            foreach (var ItemEinbauort in sEinbauortWahl)
                            {
                                progress.BeginPart(iSteps, "Export: =" + ItemAnlage + " +" + ItemEinbauort);
                                //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
                                sExportFilterSettings = "0|1|0|1120;0|0|" + ItemAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + ItemEinbauort + "|0|1|1|0|0|0;0|";
                                sExportFilename = sProjectDocPath + @"\Artikelsummenliste_" + strLabelName + "_" + ItemAnlage + "_" + ItemEinbauort + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei: 
                                ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                                iSteps += iSteps;
                                progress.EndPart();                             
                            }
                        }
                    }
                    else if ((sEinbauortWahl.Count == 0) & (sAnlagenWahl.Count > 0))
                    {

                        foreach (var ItemAnlage in sAnlagenWahl)
                        {
                            progress.BeginPart(iSteps, "Export: =" + ItemAnlage );
                            // MessageBox.Show("nur Anlage" + ItemAnlage);
                            sExportFilename = sProjectDocPath + @"\Artikelsummenliste_" + strLabelName + "_" + ItemAnlage + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:
                            sExportFilterSettings = "0|1|0|1120;0|0|" + ItemAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + "" + "|0|1|1|0|0|0;0|";
                            ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                            iSteps += iSteps;
                            progress.EndPart();
                        }
                    }
                    else if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count == 0))
                    {

                        foreach (var ItemEinbauort in sEinbauortWahl)
                        {
                            progress.BeginPart(iSteps, "Export: +" + ItemEinbauort);
                            sExportFilterSettings = "0|0|0|1120;0|0|" + "" + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + ItemEinbauort + "|0|1|1|0|0|0;0|";
                            sExportFilename = sProjectDocPath + @"\Artikelsummenliste_" + strLabelName + "_" + ItemEinbauort + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:
                            ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                            iSteps += iSteps;
                            progress.EndPart();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Fehler");
                    }
                }

            }
            finally
            {
                progress.EndPart(true);
            }
            // Meldung Erzeugen
            string sExportierteOrte;
            string sExportierteAnlagen;
            if (sEinbauortWahl.Count > 0)
            {
                sExportierteOrte = string.Join(",\r\n+", sEinbauortWahl.ToArray());
                sExportierteOrte = "Einbauorte:\r\n+" + sExportierteOrte;
            }
            else
            {
                sExportierteOrte = string.Empty;
            }
            if (sAnlagenWahl.Count > 0)
            {
                sExportierteAnlagen = string.Join(",\r\n=", sAnlagenWahl.ToArray());
                sExportierteAnlagen = "Anlagen:\r\n=" + sExportierteAnlagen;
            }
            else
            {
                sExportierteAnlagen = string.Empty;
            }

            MessageBox.Show("Artikeldaten exportiert für:\r\n" + sExportierteOrte + "\r\n" + sExportierteAnlagen);

            DialogResult Result = MessageBox.Show(
                "Sollen weitere Daten exportiert werden?",
                "Export Daten",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (Result == DialogResult.No)
                this.Close();
        }


        private static string GetProject()
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

        public void DeleteFile(string filename)
        {
            //string fileName = @"C:\test\test.txt";

            if (File.Exists(filename))
            {
                File.Delete(filename);
              //  MessageBox.Show("Datei gelöscht");
            }

            return;
        }

        private static bool ImportScheme(string schemeName, string SETTINGS_PATH, string sFilename)       
        {
            /// <summary>
            /// Import eines Eplan Schemes
            /// </summary>
            /// <remarks>
            /// string schemeName Name des Schematas
            /// string SETTINGS_PATH = wenn unklar Daten aus XML lesen Beispiel für Beschriftung: "USER.Labelling.Config";
            /// string sFilname Verweis auf XML-Datei, Beispiel: @"$(MD_SCRIPTS)\Test\LB.Einbauorte.xml"
            /// </remarks>
            string saufgeloesst = PathMap.SubstitutePath(sFilename);
            SchemeSetting schemeSetting = new SchemeSetting();
            schemeSetting.Init(SETTINGS_PATH);
            try
            {
                // prüfen ob die das Schema vorhanden ist, import wenn nicht da
                if (schemeSetting.CheckIfSchemeExists(schemeName) == false)
                {
                    // testen ob es die XML Datei gibt
                    if (FunctionCheckFileExist(saufgeloesst))
                    {
                        // import wenn Schema nicht in Eplan und XML-Datei existiert
                        schemeSetting.ImportSchemes(saufgeloesst, false);
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
                if (File.Exists(sfileName))
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

        private void ExportPartlist(string sSchemename, string SETTINGS_PATH, string sExportFilename, string strProjectName, string settingsExport, string sSortscheme)
        {                    

            using (Settings oSett = new Settings())
            {
                oSett.SetStringSetting(SETTINGS_PATH + "." + sSchemename + ".Data.SortFilter.FilterSchemeData", settingsExport, 0);
/*
                   // if (sAnlage == "nur_Einbauort")
                   // {
                   //                          nur Einbauort                                                             0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|   S1.F3          |0|1|1|0|0|0;0|
                  // oSett.SetStringSetting("USER.Labelling.Config." + sSchemename + ".Data.SortFilter.FilterSchemeData", "0|0|0|1120;0|0|" + sAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sEinbauort + "|0|1|1|0|0|0;0|", 0);
                    // oSett.SetStringSetting("USER.Labelling.Config." + sSchemename + ".Data.SortFilter.FilterSchemeName", "Einbauort");
               // }
               // else if ((sEinbauort == "nur_Anlage"))
               // {
                    //                          nur Anlage                                                                0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0| S1.F3 |0|1|1|0|0|0;0|
               //     oSett.SetStringSetting("USER.Labelling.Config." + sSchemename + ".Data.SortFilter.FilterSchemeData", "0|1|0|1120;0|0|" + sAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + sEinbauort + "|0|1|1|0|0|0;0|", 0);    
               // }
                // Einbauort und Anlage
               // else
                //{         
                //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
               // oSett.SetStringSetting("USER.Labelling.Config." + sSchemename + ".Data.SortFilter.FilterSchemeData", "0|1|0|1120;0|0|" + sAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sEinbauort + "|0|1|1|0|0|0;0|", 0);
                    // oSett.SetStringSetting("USER.Labelling.Config." + sSchemename + ".Data.SortFilter.FilterSchemeName", "Anlage und Einbauort");
               // }
  */
    }

            // Parameter für die Action
            ActionCallingContext labellingContext = new ActionCallingContext();

            labellingContext.AddParameter("PROJECTNAME", strProjectName); // Projektname komplett mit Erweiterung
            labellingContext.AddParameter("CONFIGSCHEME", sSchemename);//"ExportFunc_Artikelsummenstueckliste");
            labellingContext.AddParameter("FILTERSCHEME", ""); // Hier darf kein Filterschema eingetragen werden
            labellingContext.AddParameter("SORTSCHEME", sSortscheme); // Sortieren
            labellingContext.AddParameter("LANGUAGE", "de_DE");
            labellingContext.AddParameter("DESTINATIONFILE", sExportFilename); // sProjectDocPath + @"\Artikelsummen" + "_" + strLabelName + "_" + sAnlage + "_" + sEinbauort + "_" + myDate + "_" + myTime + ".xlsx"); // Zieldatei: 
            labellingContext.AddParameter("RECREPEAT", "1");
            labellingContext.AddParameter("TASKREPEAT", "1");
            labellingContext.AddParameter("SHOWOUTPUT", "0");

            // Ausführen der Beschriftungsaction mit Parametern
            new CommandLineInterpreter().Execute("label", labellingContext);
        }


        private static void LabellingExport(string sfilename, string strProjectName, string sScheme)
        {
            // Ausgabe der Einbauorte 
            // Parameter für die Action
            ActionCallingContext labellingContext = new ActionCallingContext();
            labellingContext.AddParameter("PROJECTNAME", strProjectName); // Projektname komplett mit Erweiterung
            labellingContext.AddParameter("CONFIGSCHEME", sScheme);
            labellingContext.AddParameter("FILTERSCHEME", ""); // Hier darf kein Filterschema eingetragen werden
            labellingContext.AddParameter("SORTSCHEME", "");
            labellingContext.AddParameter("LANGUAGE", "de_DE");
            labellingContext.AddParameter("DESTINATIONFILE", sfilename); // Zieldatei: "Projektpfad/Projektname_Testbeschriftung.xls"
            labellingContext.AddParameter("RECREPEAT", "1");
            labellingContext.AddParameter("TASKREPEAT", "1");
            labellingContext.AddParameter("SHOWOUTPUT", "0");
            labellingContext.AddParameter("USESELECTION", "0");   //Auswahl berücksichtigen: 0 = Nein, 1 = Ja

            // Ausführen der Beschriftungsaction mit Parametern
            new CommandLineInterpreter().Execute("label", labellingContext);
        }

        private void LoadFrmSelect(object sender, EventArgs e)
        {
            /// <summary>
            /// wird aufgerufen wenn deas Formular geöffnet wurde
            /// </summary>

            string strProjectNameFull = GetProject();
            // Dokumentenordner auswählen
            string sProjectDocPath = PathMap.SubstitutePath("$(DOC)");
            string sfilenameEinbauort = sProjectDocPath + @"\Einbauorte_" + myDate + "_" + myTime + ".txt";
            string sfilenameAnlagen = sProjectDocPath + @"\Anlagen_" + myDate + "_" + myTime + ".txt";
            ///<remarks>
            /// Variablen für den Import des Schemas Einbauorte
            /// </remarks>
            string schemeNameEinbauorte = "ExportFunc_Einbauorte";
            string schemeNameAnlagen = "ExportFunc_Anlagen";
            string SETTINGS_PATH = "USER.Labelling.Config";
            string sImportFileEinbauorte = @"$(MD_SCRIPTS)\Ausgabe\LB.ExportFunc_Einbauorte.xml";
            string sImportFileAnlagen = @"$(MD_SCRIPTS)\Ausgabe\LB.ExportFunc_Anlagen.xml";
            ///<remarks>
            /// EplanMacro mit Pfad zum Script Ordner auflösen
            ///<\remarks>

            ///<remarks>
            /// Import des Schemas für die Einbauorte
            ///<\remarks>
            bool result = ImportScheme(schemeNameEinbauorte, SETTINGS_PATH, sImportFileEinbauorte);
            if (result == false)
            {
                MessageBox.Show("Fehler beim Import");
            }
            ///<remarks>
            /// Ausgabe der TXT-Datei mit allen Einbauorten im Projekt
            ///</remarks>
            LabellingExport(sfilenameEinbauort, strProjectNameFull, schemeNameEinbauorte);
            ///<remarks>
            /// Import des Schemas für die Anlagen
            ///<\remarks>
           
            result = ImportScheme(schemeNameAnlagen, SETTINGS_PATH, sImportFileAnlagen);
            if (result == false)
            {
                MessageBox.Show("Fehler beim Import");
            }
            ///<remarks>
            /// Ausgabe der TXT-Datei mit allen Anlagen im Projekt
            ///</remarks>
            LabellingExport(sfilenameAnlagen, strProjectNameFull, schemeNameAnlagen);

            ///<remarks>
            /// Textdatei  mit Einbauorten lesen
            ///</remarks>
             List<string> sEinbauorteAll = new List<string>();
            sEinbauorteAll = ReadFile(sfilenameEinbauort);
           

            ///<remarks>
            /// Textdatei  mit Einbauorten lesen
            ///</remarks>
            List<string> sAnlagenAll = new List<string>();
            sAnlagenAll = ReadFile(sfilenameAnlagen);
            ///<remarks>
            ///ermitteln der maximalen Größe
            ///</remarks>

            int iSizeXa = 20 * sAnlagenAll.Count(); // + 40;
            int iSizeXl = 20 * sEinbauorteAll.Count();// + 40;

            int iSize;

            if (iSizeXa > iSizeXl)
                iSize = iSizeXa;
            else
                iSize = iSizeXl;

            iSize = (iSize / 3) + 40; // Daten werden in drei Reihen dargestellt

           // Größe des Formulares anpassen
            this.ClientSize = new System.Drawing.Size(1150, iSize);
            ///<remarks>
            /// Checkboxen dynamisch erzeugen
            ///</remarks>
            CreateBoxString(sAnlagenAll, sTagAnlagen, 0, iSize);
            ///<remarks>
            /// Checkboxen dynamisch erzeugen
            ///</remarks>
            CreateBoxString(sEinbauorteAll, sTagEinbauort, 550, iSize);
            ///<remarks>
            ///löschen der überflüssigen Dateien
            ///</remarks> 
            DeleteFile(sfilenameAnlagen);
            DeleteFile(sfilenameEinbauort);
            this.Text = "Ausgabe von Listen für: " + strLabelName;
        }

        private void Box_SummaryChanged(object objSender, EventArgs e)
        {
            ///<summary>
            /// Abfrage ob Summenliste erzeugt werden soll
            /// </summary>
            
            CheckBox cb = objSender as CheckBox; // get a reference to the checked/unchecked CheckBox

            // Abfragen ob gesetzt
            if (cb.Checked)
            {
               // MessageBox.Show("gesetzt");
                getSummaryList = true;
            }
            else
            {
               // MessageBox.Show("zurück gesetzt");
                    getSummaryList = false;
            }
        }

        private void ButtonPartlist_Click(object sender, EventArgs e)
        {

            string sSchemename = "ExportFunc_Artikelstueckliste";
            string SETTINGS_PATH = "USER.Labelling.Config";
            string sImportFileArtikelstueckliste = @"$(MD_SCRIPTS)\Ausgabe\LB.ExportFunc_Artikelstueckliste.xml";

            bool result = ImportScheme(sSchemename, SETTINGS_PATH, sImportFileArtikelstueckliste);

            if (!result)
            {
                MessageBox.Show("Fehler beim Import der Filter");
                return;
            }


            Progress progress = new Progress("EnhancedProgress");
            progress.SetAllowCancel(false);
            progress.ShowImmediately();
            progress.SetTitle("Ausgabe Stücklisten");

            try
            {

                string sSummaryLocation;
                string sSummaryHigherLevel;
                string sExportFilterSettings;
                string sExportFilename;
                string sSortscheme = "Anlage+Einbauort+BMK"; 
                // Prüfen ob eine Wahl erfolgt ist
                if ((sEinbauortWahl.Count == 0) & (sAnlagenWahl.Count == 0))
                {
                    MessageBox.Show("Bitte einen Einbauort oder eine Anlage wählen");
                    return;
                }

                // Ausgabe Stückliste
                if (getSummaryList)
                {
                    progress.BeginPart(33, "Stückliste");
                    // aus der Liste einen String machen
                    sSummaryHigherLevel = string.Join(";", sAnlagenWahl.ToArray());
                    sSummaryLocation = string.Join(";", sEinbauortWahl.ToArray());

                    if (!(string.IsNullOrEmpty(sSummaryHigherLevel)) & !(string.IsNullOrEmpty(sSummaryLocation)))
                    {
                        //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
                       // MessageBox.Show("Anlage und Einbauort" + sSummaryHigherLevel + sSummaryLocation);
                        sExportFilterSettings = "0|1|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Artikelstückliste" + "_" + strLabelName + "_" + sSummaryHigherLevel.Replace(";", "_") + "_" + sSummaryLocation.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei: 
                    }
                    else if (string.IsNullOrEmpty(sSummaryLocation))
                    {
                        //                          nur Anlage                                                                0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0| S1.F3 |0|1|1|0|0|0;0|
                       // MessageBox.Show("Anlage");
                        sExportFilterSettings = "0|1|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Artikelstückliste" + "_" + strLabelName + "_" + sSummaryHigherLevel.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:
                    }
                    else //(string.IsNullOrEmpty(sSummaryHigherLevel))
                    {
                        //                          nur Einbauort                                                             0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|   S1.F3          |0|1|1|0|0|0;0|
                       // MessageBox.Show("nur Einbauort");
                        sExportFilterSettings = "0|0|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Artikelstückliste" + "_" + strLabelName + "_" + sSummaryLocation.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:

                    }

                    ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                    progress.EndPart();
                }
                else
                {
                    // Progressbar zusammen bauen
                    double iSteps = 1;

                    if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count > 0))
                        iSteps = 1 / (sEinbauortWahl.Count * sAnlagenWahl.Count);
                    else if (sAnlagenWahl.Count > 0)
                    {
                        iSteps = 1 / sAnlagenWahl.Count;
                    }
                    else if (sEinbauortWahl.Count > 0)
                    {
                        iSteps = 1 / sEinbauortWahl.Count;
                    }

                    // Ausgabe der Daten in einzelne Listen

                    if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count > 0))
                    {
                        foreach (var ItemAnlage in sAnlagenWahl)
                        {
                            foreach (var ItemEinbauort in sEinbauortWahl)
                            {
                                progress.BeginPart(iSteps, "Export: =" + ItemAnlage + " +" + ItemEinbauort);
                                //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
                                sExportFilterSettings = "0|1|0|1120;0|0|" + ItemAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + ItemEinbauort + "|0|1|1|0|0|0;0|";
                                sExportFilename = sProjectDocPath + @"\Artikelstückliste" + "_" + strLabelName + "_" + ItemAnlage + "_" + ItemEinbauort + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei: 
                                ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                                iSteps += iSteps;
                                progress.EndPart();

                            }
                        }
                    }
                    else if ((sEinbauortWahl.Count == 0) & (sAnlagenWahl.Count > 0))
                    {

                        foreach (var ItemAnlage in sAnlagenWahl)
                        {
                            progress.BeginPart(iSteps, "Export: =" + ItemAnlage + " +" + ItemAnlage);
                            sExportFilename = sProjectDocPath + @"\Artikelstückliste" + "_" + strLabelName + "_" + ItemAnlage + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:
                            sExportFilterSettings = "0|1|0|1120;0|0|" + ItemAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + "" + "|0|1|1|0|0|0;0|";
                            ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                            iSteps += iSteps;
                            progress.EndPart();
                        }
                    }
                    else if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count == 0))
                    {

                        foreach (var ItemEinbauort in sEinbauortWahl)
                        {
                            progress.BeginPart(iSteps, "Export: +" + ItemEinbauort);
                            sExportFilterSettings = "0|0|0|1120;0|0|" + "" + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + ItemEinbauort + "|0|1|1|0|0|0;0|";
                            sExportFilename = sProjectDocPath + @"\Artikelstückliste" + "_" + strLabelName + "_" + ItemEinbauort + "_" + myDate + "_" + myTime + ".xlsx"; // Zieldatei:
                            ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                            iSteps += iSteps;
                            progress.EndPart();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Fehler");
                    }
                }

            }
            finally
            {
                progress.EndPart(true);
            }
            // Meldung Erzeugen
            string sExportierteOrte;
            string sExportierteAnlagen;
            if (sEinbauortWahl.Count > 0)
            {
                sExportierteOrte = string.Join(",\r\n+", sEinbauortWahl.ToArray());
                sExportierteOrte = "Einbauorte:\r\n+" + sExportierteOrte;
            }
            else
            {
                sExportierteOrte = string.Empty;
            }
            if (sAnlagenWahl.Count > 0)
            {
                sExportierteAnlagen = string.Join(",\r\n=", sAnlagenWahl.ToArray());
                sExportierteAnlagen = "Anlagen:\r\n=" + sExportierteAnlagen;
            }
            else
            {
                sExportierteAnlagen = string.Empty;
            }

            MessageBox.Show("Artikeldaten exportiert für:\r\n" + sExportierteOrte + "\r\n" + sExportierteAnlagen);

            DialogResult Result = MessageBox.Show(
           "Sollen weitere Daten exportiert werden?",
           "Export Daten",
           MessageBoxButtons.YesNo,
           MessageBoxIcon.Question
           );

            if (Result == DialogResult.No)
                this.Close();

        }

        private void ButtonCablelist_Click(object sender, EventArgs e)
        {

            string sSchemename = "ExportFunc_Kabelliste";
            string SETTINGS_PATH = "USER.Labelling.Config";
            string sImportFileScheme = @"$(MD_SCRIPTS)\Ausgabe\LB.ExportFunc_Kabelliste.xml";

            bool result = ImportScheme(sSchemename, SETTINGS_PATH, sImportFileScheme);

            if (!result)
            {
                MessageBox.Show("Fehler beim Import der Filter");
                return;
            }


            Progress progress = new Progress("EnhancedProgress");
            progress.SetAllowCancel(false);
            progress.ShowImmediately();
            progress.SetTitle("Ausgabe Kabellisten");

            try
            {

                string sSummaryLocation;
                string sSummaryHigherLevel;
                string sExportFilterSettings;
                string sExportFilename;
                string sSortscheme = "";
                // Prüfen ob eine Wahl erfolgt ist
                if ((sEinbauortWahl.Count == 0) & (sAnlagenWahl.Count == 0))
                {
                    MessageBox.Show("Bitte einen Einbauort oder eine Anlage wählen");
                    return;
                }

                // Ausgabe Stückliste
                if (getSummaryList)
                {
                    progress.BeginPart(33, "Stückliste");
                    // aus der Liste einen String machen
                    sSummaryHigherLevel = string.Join(";", sAnlagenWahl.ToArray());
                    sSummaryLocation = string.Join(";", sEinbauortWahl.ToArray());

                    if (!(string.IsNullOrEmpty(sSummaryHigherLevel)) & !(string.IsNullOrEmpty(sSummaryLocation)))
                    {
                        //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
                        // MessageBox.Show("Anlage und Einbauort" + sSummaryHigherLevel + sSummaryLocation);
                        sExportFilterSettings = "0|1|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Kabelliste" + "_" + strLabelName + "_" + sSummaryHigherLevel.Replace(";", "_") + "_" + sSummaryLocation.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsm"; // Zieldatei: 
                    }
                    else if (string.IsNullOrEmpty(sSummaryLocation))
                    {
                        //                          nur Anlage                                                                0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0| S1.F3 |0|1|1|0|0|0;0|
                        // MessageBox.Show("Anlage");
                        sExportFilterSettings = "0|1|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Kabelliste" + "_" + strLabelName + "_" + sSummaryHigherLevel.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsm"; // Zieldatei:
                    }
                    else //(string.IsNullOrEmpty(sSummaryHigherLevel))
                    {
                        //                          nur Einbauort                                                             0|1|0|1120;0|0| MSR; HV1      |0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|   S1.F3          |0|1|1|0|0|0;0|
                        // MessageBox.Show("nur Einbauort");
                        sExportFilterSettings = "0|0|0|1120;0|0|" + sSummaryHigherLevel + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + sSummaryLocation + "|0|1|1|0|0|0;0|";
                        sExportFilename = sProjectDocPath + @"\Kabelliste" + "_" + strLabelName + "_" + sSummaryLocation.Replace(";", "_") + "_" + myDate + "_" + myTime + ".xlsm"; // Zieldatei:

                    }

                    ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                    progress.EndPart();
                }
                else
                {
                    // Progressbar zusammen bauen
                    double iSteps = 1;

                    if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count > 0))
                        iSteps = 1 / (sEinbauortWahl.Count * sAnlagenWahl.Count);
                    else if (sAnlagenWahl.Count > 0)
                    {
                        iSteps = 1 / sAnlagenWahl.Count;
                    }
                    else if (sEinbauortWahl.Count > 0)
                    {
                        iSteps = 1 / sEinbauortWahl.Count;
                    }

                    // Ausgabe der Daten in einzelne Listen

                    if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count > 0))
                    {
                        foreach (var ItemAnlage in sAnlagenWahl)
                        {
                            foreach (var ItemEinbauort in sEinbauortWahl)
                            {
                                progress.BeginPart(iSteps, "Export: =" + ItemAnlage + " +" + ItemEinbauort);
                                //                              Anlage und Einbauort                                                  0|1|0|1120;0|0|    EB3       |0|1|1|0|0|0;0|#7|1|0|;0|0| |0|1|1|0|0|0;0|#0|1|0|1220;0|0|   ET1      |0|1|1|0|0|0;0|                   
                                sExportFilterSettings = "0|1|0|1120;0|0|" + ItemAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + ItemEinbauort + "|0|1|1|0|0|0;0|";
                                sExportFilename = sProjectDocPath + @"\Kabelliste" + "_" + strLabelName + "_" + ItemAnlage + "_" + ItemEinbauort + "_" + myDate + "_" + myTime + ".xlsm"; // Zieldatei: 
                                ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                                iSteps += iSteps;
                                progress.EndPart();

                            }
                        }
                    }
                    else if ((sEinbauortWahl.Count == 0) & (sAnlagenWahl.Count > 0))
                    {

                        foreach (var ItemAnlage in sAnlagenWahl)
                        {
                            progress.BeginPart(iSteps, "Export: =" + ItemAnlage + " +" + ItemAnlage);
                            sExportFilename = sProjectDocPath + @"\Kabelliste" + "_" + strLabelName + "_" + ItemAnlage + "_" + myDate + "_" + myTime + ".xlsm"; // Zieldatei:
                            sExportFilterSettings = "0|1|0|1120;0|0|" + ItemAnlage + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|1220;0|0|" + "" + "|0|1|1|0|0|0;0|";
                            ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                            iSteps += iSteps;
                            progress.EndPart();
                        }
                    }
                    else if ((sEinbauortWahl.Count > 0) & (sAnlagenWahl.Count == 0))
                    {

                        foreach (var ItemEinbauort in sEinbauortWahl)
                        {
                            progress.BeginPart(iSteps, "Export: +" + ItemEinbauort);
                            sExportFilterSettings = "0|0|0|1120;0|0|" + "" + "|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|1220;0|0|" + ItemEinbauort + "|0|1|1|0|0|0;0|";
                            sExportFilename = sProjectDocPath + @"\Kabelliste" + "_" + strLabelName + "_" + ItemEinbauort + "_" + myDate + "_" + myTime + ".xlsm"; // Zieldatei:
                            ExportPartlist(sSchemename, SETTINGS_PATH, sExportFilename, GetProject(), sExportFilterSettings, sSortscheme);
                            iSteps += iSteps;
                            progress.EndPart();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Fehler");
                    }
                }

            }
            finally
            {
                progress.EndPart(true);
            }
            // Meldung Erzeugen
            string sExportierteOrte;
            string sExportierteAnlagen;
            if (sEinbauortWahl.Count > 0)
            {
                sExportierteOrte = string.Join(",\r\n+", sEinbauortWahl.ToArray());
                sExportierteOrte = "Einbauorte:\r\n+" + sExportierteOrte;
            }
            else
            {
                sExportierteOrte = string.Empty;
            }
            if (sAnlagenWahl.Count > 0)
            {
                sExportierteAnlagen = string.Join(",\r\n=", sAnlagenWahl.ToArray());
                sExportierteAnlagen = "Anlagen:\r\n=" + sExportierteAnlagen;
            }
            else
            {
                sExportierteAnlagen = string.Empty;
            }

            MessageBox.Show("Kabellisten exportiert für:\r\n" + sExportierteOrte + "\r\n" + sExportierteAnlagen);

            DialogResult Result = MessageBox.Show(
           "Sollen weitere Daten exportiert werden?",
           "Export Daten",
           MessageBoxButtons.YesNo,
           MessageBoxIcon.Question
           );

            if (Result == DialogResult.No)
                this.Close();
        }
    }
}
