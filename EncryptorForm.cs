using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace outlookaddin
{
    public partial class EncryptorForm : Form
    {
        private String tempMessage = null;
        private bool cryption = false;
        public EncryptorForm()
        {
            InitializeComponent();
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            String pass = textBox1.Text;
            if(pass != ThisAddIn.getPassword())
            {
                ThisAddIn.setPassword(pass);
            }

            Tuple<String,String> cryptedtuple = Encrypt.EncryptString(OutlookHelperTool.getMailText(false), pass);
            String crypted = cryptedtuple.Item1;
            String vector = cryptedtuple.Item2;

            
            tempMessage = OutlookHelperTool.getPreparedMessage(crypted, vector);
           
            cryption = true;
            //OutlookHelperTool.setMailText(exportString);
            buttonApply.Enabled = true;
            buttonPreview.Enabled = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            String pass = textBox1.Text;
            if (pass != ThisAddIn.getPassword())
            {
                ThisAddIn.setPassword(pass);
            }


            String text = OutlookHelperTool.getMailText(true);
            tempMessage = Encrypt.DecryptString(text, pass, OutlookHelperTool.getinitVector(OutlookHelperTool.getMailText(false)));
            
            //OutlookHelperTool.setMailText(Encrypt.DecryptString(text, textBox1.Text,OutlookHelperTool.getinitVector(text)));
            buttonApply.Enabled = true;
            buttonPreview.Enabled = true;
            cryption = false;


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            buttonApply.Enabled = false;
            buttonPreview.Enabled = false;
            OutlookHelperTool.setMailText(tempMessage);
            if (cryption)
            {
                OutlookHelperTool.setCryptedSubject();
            }
            else
            {
                OutlookHelperTool.cutCryptedSubject();
            }
            this.Close();
        }

        private void buttonPreview_Click(object sender, EventArgs e)
        {
            Preview preview = new Preview(tempMessage);
            buttonPreview.Enabled = false;

        }

        private void EncryptorForm_Load(object sender, EventArgs e)
        {
            if (ThisAddIn.usePassword)
            {
                textBox1.Text = ThisAddIn.getPassword();
            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }
    }
}
