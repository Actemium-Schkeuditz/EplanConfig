using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO;
using System.Collections.Generic;

public class TestAusgabeEinbauorte
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


        string sfilenameEinbauort = sProjectDocPath + @"\Einbauorte_" + myDate + "_" + myTime + ".xml";
        // test
        //  MessageBox.Show(sfilenameEinbauort);

        // Parameter für die Action
        ActionCallingContext labellingContext = new ActionCallingContext();
        labellingContext.AddParameter("PROJECTNAME", strProjectName); // Projektname komplett mit Erweiterung
        labellingContext.AddParameter("CONFIGSCHEME", "Einbauorte");
        labellingContext.AddParameter("FILTERSCHEME", ""); // Hier darf kein Filterschema eingetragen werden
        labellingContext.AddParameter("SORTSCHEME", "");
        labellingContext.AddParameter("LANGUAGE", "de_DE");
        labellingContext.AddParameter("DESTINATIONFILE", sfilenameEinbauort); // Zieldatei: "Projektpfad/Projektname_Testbeschriftung.xls"
        labellingContext.AddParameter("RECREPEAT", "1");
        labellingContext.AddParameter("TASKREPEAT", "1");
        labellingContext.AddParameter("SHOWOUTPUT", "0");
        labellingContext.AddParameter("USESELECTION", "0");   //auswahlh berücksichtigen: 0 = Nein, 1 = Ja

        // Ausführen der Beschriftungsaction mit Parametern
        new CommandLineInterpreter().Execute("label", labellingContext);


        // lesen der XML-Datei mit den Einbauorten

        List<EplanLabellingDocumentPageLineLabelProperty> persons = ReadXml(sfilenameEinbauort);
        foreach (EplanLabellingDocumentPageLineLabelProperty person in persons)
        {
            MessageBox.Show(person.PropertyValue);
        }

        return;


    }

    // Menü zusammenbauen
    [DeclareMenu]
    public void MenuFunction()
    {
        //Deklarationen
        //uint MenuID = new uint(); // Menü-ID vom neu erzeugten Menü   

        Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();

        uint MenuID = oMenu.GetPersistentMenuId("Export / Beschriftung...");

        MessageBox.Show(MenuID.ToString());

        oMenu.AddMenuItem(
                "Einbauorte", // Name: Menüpunkt
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


    public static List<EplanLabellingDocumentPageLineLabelProperty> ReadXml(string fileName)
    {
        XmlSerializer serializer = new XmlSerializer(typeof(List<EplanLabellingDocumentPageLineLabelProperty>));
        StreamReader reader = new StreamReader(fileName);

        List<EplanLabellingDocumentPageLineLabelProperty> StrukturKennzeichen = (List<EplanLabellingDocumentPageLineLabelProperty>)serializer.Deserialize(reader);
        reader.Close();
        reader.Dispose();
        return StrukturKennzeichen;
    }


//------------------------------------------------------------------------------
// <auto-generated>
//     Dieser Code wurde von einem Tool generiert.
//     Laufzeitversion:4.0.30319.42000
//
//     Änderungen an dieser Datei können falsches Verhalten verursachen und gehen verloren, wenn
//     der Code erneut generiert wird.
// </auto-generated>
//------------------------------------------------------------------------------



// 
// Dieser Quellcode wurde automatisch generiert von xsd, Version=4.6.1055.0.
// 


/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
[System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
public partial class EplanLabelling
{

    private EplanLabellingDocument documentField;

    private decimal versionField;

    /// <remarks/>
    public EplanLabellingDocument Document
    {
        get => this.documentField;
        set => this.documentField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public decimal version
    {
        get => this.versionField;
        set => this.versionField = value;
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class EplanLabellingDocument
{

    private EplanLabellingDocumentPage pageField;

    private byte source_idField;

    /// <remarks/>
    public EplanLabellingDocumentPage Page
    {
        get => this.pageField;
        set => this.pageField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte source_id
    {
        get => this.source_idField;
        set => this.source_idField = value;
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class EplanLabellingDocumentPage
{

    private object headerField;

    private EplanLabellingDocumentPageColumnHeader columnHeaderField;

    private EplanLabellingDocumentPageLine[] lineField;

    private object footerField;

    private byte source_idField;

    /// <remarks/>
    public object Header
    {
        get => this.headerField;
        set => this.headerField = value;
    }

    /// <remarks/>
    public EplanLabellingDocumentPageColumnHeader ColumnHeader
    {
        get => this.columnHeaderField;
        set => this.columnHeaderField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute("Line")]
    public EplanLabellingDocumentPageLine[] Line
    {
        get => this.lineField;
        set => this.lineField = value;
    }

    /// <remarks/>
    public object Footer
    {
        get => this.footerField;
        set => this.footerField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte source_id
    {
        get => this.source_idField;
        set => this.source_idField = value;
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class EplanLabellingDocumentPageColumnHeader
{

    private string propertyNameField;

    private string dataTypeField;

    /// <remarks/>
    public string PropertyName
    {
        get => this.propertyNameField;
        set => this.propertyNameField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string DataType
    {
        get => this.dataTypeField;
        set => this.dataTypeField = value;
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class EplanLabellingDocumentPageLine
{

    private EplanLabellingDocumentPageLineLabel labelField;

    private byte source_idField;

    private string separatorField;

    /// <remarks/>
    public EplanLabellingDocumentPageLineLabel Label
    {
        get => this.labelField;
        set => this.labelField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte source_id
    {
        get => this.source_idField;
        set => this.source_idField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string separator
    {
        get => this.separatorField;
        set => this.separatorField = value;
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class EplanLabellingDocumentPageLineLabel
{

    private EplanLabellingDocumentPageLineLabelProperty propertyField;

    private byte source_idField;

    /// <remarks/>
    public EplanLabellingDocumentPageLineLabelProperty Property
    {
        get => this.propertyField;
        set => this.propertyField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte source_id
    {
        get => this.source_idField;
        set => this.source_idField = value;
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class EplanLabellingDocumentPageLineLabelProperty
{

    private string propertyNameField;

    private string propertyValueField;

    private byte formattingTypeField;

    private byte formattingLengthField;

    private byte formattingRAlignField;

    /// <remarks/>
    public string PropertyName
    {
        get => this.propertyNameField;
        set => this.propertyNameField = value;
    }

    /// <remarks/>
    public string PropertyValue
    {
        get => this.propertyValueField;
        set => this.propertyValueField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte FormattingType
    {
        get => this.formattingTypeField;
        set => this.formattingTypeField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte FormattingLength
    {
        get => this.formattingLengthField;
        set => this.formattingLengthField = value;
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public byte FormattingRAlign
    {
        get => this.formattingRAlignField;
        set => this.formattingRAlignField = value;
    }
}
}
