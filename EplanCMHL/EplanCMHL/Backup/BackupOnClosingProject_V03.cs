// V2 24.10.2017
// Christian Langrock
//Änderung Datumsformat
// 26.08.2019 CL
//Abfrage ob die richtigen Domain Nutzer drin sind 
// das Skript wird nur noch ausgeführt wenn die Nutzer in der Gruppe VED\ORG-AC-DESK001-BU00103-Users drin sind
// Änderung zu einem Script welches als Event von einem anderen Script aufgerufen wird

using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using MessageBox = System.Windows.Forms.MessageBox;


public class BackupOnClosingProject
{
    //[DeclareEventHandler("onActionStart.String.XPrjActionProjectClose")]
    [Start]
    public void Function()
    {
        string strProjectname = PathMap.SubstitutePath("$(PROJECTNAME)");
		string strFullProjectname = PathMap.SubstitutePath("$(P)");

		string sProjectpath = PathMap.SubstitutePath("$(PROJECTPATH)");
		sProjectpath = sProjectpath.Substring(0, sProjectpath.LastIndexOf("\\"));

		//Sicherungsverzeichnis auswählen
		sProjectpath += "\\SICHERUNGEN";
        //string sBackupPath = PathMap.SubstitutePath(@"$(EPLAN_DATA)\..\41_BAK\deSK001\Projekte\CMHL");

       // MessageBox.Show(sBackupPath);


            DialogResult Result = MessageBox.Show(
           "Soll eine Sicherung fuer das Projekt\n'"
           + strProjectname +
           "'\nerzeugt werden?\n"
           + "\nSicherungsverzeichnis:\n" + sProjectpath,
           "Datensicherung",
           MessageBoxButtons.YesNo,
           MessageBoxIcon.Question
           );

            if (Result == DialogResult.Yes)

            {
                string myTime = System.DateTime.Now.ToString("yyyyMMdd");
                string hour = System.DateTime.Now.Hour.ToString();
                string minute = System.DateTime.Now.Minute.ToString();

                Progress progress = new Progress("SimpleProgress");
                progress.BeginPart(100, "");
                progress.SetAllowCancel(true);

                if (!progress.Canceled())
                {
                    progress.BeginPart(33, "Backup");
                    ActionCallingContext backupContext = new ActionCallingContext();
                    backupContext.AddParameter("BACKUPMEDIA", "DISK");
                    backupContext.AddParameter("BACKUPMETHOD", "BACKUP");
                    backupContext.AddParameter("COMPRESSPRJ", "0");
                    backupContext.AddParameter("INCLEXTDOCS", "1");
                    backupContext.AddParameter("BACKUPAMOUNT", "BACKUPAMOUNT_ALL");
                    backupContext.AddParameter("INCLIMAGES", "1");
                    backupContext.AddParameter("destinationpath", sProjectpath);
                    backupContext.AddParameter("PROJECTNAME", strFullProjectname);
                    backupContext.AddParameter("TYPE", "PROJECT");
                    backupContext.AddParameter("ARCHIVENAME", strProjectname + "_" + myTime + "_" + hour + minute + ".");
                    new CommandLineInterpreter().Execute("backup", backupContext);
                    progress.EndPart();

                }
                progress.EndPart(true);
            }
            return;

       
    }
   
}
