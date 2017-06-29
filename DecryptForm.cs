using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace outlookaddin
{
    public partial class DecryptForm : Form
    {
        String tempMessage;

        private Outlook.MailItem mailitem;
        public DecryptForm(Outlook.MailItem mailitem)
        {
            this.mailitem = mailitem;
            InitializeComponent();
            if (ThisAddIn.usePassword)
            {
                textBox1.Text = ThisAddIn.getPassword();
            }
            this.ShowDialog();
            
        }

        private void buttonPreview_Click(object sender, EventArgs e)
        {
            Preview preview = new Preview(tempMessage);
            buttonPreview.Enabled = false;
        }

        private void buttonDecrypt_Click(object sender, EventArgs e)
        {
            String pass = textBox1.Text;
            if (pass != ThisAddIn.getPassword())
            {
                ThisAddIn.setPassword(pass);
            }
            String text = OutlookHelperTool.getMailText(mailitem,true);
            tempMessage = Encrypt.DecryptString(text, pass, OutlookHelperTool.getinitVector(OutlookHelperTool.getMailText(mailitem,false)));

            //OutlookHelperTool.setMailText(Encrypt.DecryptString(text, textBox1.Text,OutlookHelperTool.getinitVector(text)));
            buttonApply.Enabled = true;
            buttonPreview.Enabled = true;
          
        }

        private void buttonApply_Click(object sender, EventArgs e)
        {
            buttonApply.Enabled = false;
            buttonPreview.Enabled = false;
            OutlookHelperTool.setMailText(mailitem,tempMessage);
            OutlookHelperTool.cutCryptedSubject(mailitem);
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
