using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace outlookaddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            checkBoxSavePassword.Checked = ThisAddIn.usePassword;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            
            ThisAddIn.showMainForm();
        }

        private void checkBoxSavePassword_Click(object sender, RibbonControlEventArgs e)
        {
            
            ThisAddIn.setusepassword(checkBoxSavePassword.Checked);
            ThisAddIn.blankPassword();
        }
    }
}
