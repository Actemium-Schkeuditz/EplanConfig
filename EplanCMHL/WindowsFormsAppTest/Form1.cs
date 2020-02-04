using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsAppTest
{
    public partial class Form1 : Form
    {
       //private Dictionary<int, List<string>> dictDaten = new Dictionary<int, List<string>>();
        private Dictionary<string, bool> dictDaten = new Dictionary<string, bool>();

        private List<string> sEinbauortWahl = new List<String>();

        public Form1()
        {
            InitializeComponent();


            List<string> sEinbauort = new List<String>();
           
            sEinbauort.Add("nummer 1");
            sEinbauort.Add("nummer 2");
            sEinbauort.Add("nummer 3");
            
          //  dictDaten.Add(1, sEinbauort);
          //  dictDaten.Add(2, sEinbauort);
            CreateBoxString(sEinbauort);
        }


        private void CreateBoxString( List<String> daten)
        {
            //int y = 0;
            //int z = 0;
            CheckBox box;
            for (int x = 0; x < daten.Count(); x++)
            {
               // List<string> folder = dictionary[x];
                box = new CheckBox();
                box.CheckedChanged += Box_CheckedChanged; // here
                box.Text = daten[x].ToString();
                box.AutoSize = true;
                box.Location = new Point(15, ((x*20)+30));
                box.Checked = false;
                Controls.Add(box);
            }
        }

        /*
        private void CreateBox(Dictionary<int, List<String>> dictionary)
        {
            int y = 0;
            int z = 0;
            CheckBox box;
            for (int x = 1; x < dictionary.Count(); x++)
            {
                List<string> folder = dictionary[x];
                box = new CheckBox();
                box.CheckedChanged += Box_CheckedChanged; // here
                box.Text = folder[0];
                box.AutoSize = true;
              //  box.Checked = folder[1];
                Controls.Add(box);
            }
        }
        */
        private void Box_CheckedChanged(object objSender, EventArgs e)
        {
            // some code here.
            CheckBox cb = objSender as CheckBox; // get a reference to the checked/unchecked CheckBox

            // Do something just for demo
            if (cb.Checked)
            {
                //MessageBox.Show(cb.Text + " checked");
               // dictDaten.Add(cb.Text, true);
                sEinbauortWahl.Add(cb.Text);
            }
            else
            {
               // MessageBox.Show(cb.Text + " unchecked");
              //  dictDaten.Remove(cb.Text);
                sEinbauortWahl.Remove(cb.Text);
            }
        }

        private void ButtonCancel_Click(object sender, EventArgs e)
        {


            
            


            this.Close();
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {

            /* StringBuilder sb = new StringBuilder();
           foreach (var item in dictDaten)
             {
                 sb.AppendFormat("{0} - {1}{2}", item.Key, item.Value, Environment.NewLine);
             }
             string result = sb.ToString().TrimEnd();//when converting to string we also want to trim the redundant new line at the very end
              MessageBox.Show(result);
              }*/

            


            string result = string.Join("\r\n",sEinbauortWahl.ToArray() );

            result = result.ToString().TrimStart();//when converting to string we also want to trim the redundant new line at the very end
            MessageBox.Show(result);

            // hier fehlt noch Aktion
        }

        private void ListViewResult_DoubleClick(object sender, EventArgs e)
        {
            // hier ist was offen
        }

        private void ListViewResult_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }

}
