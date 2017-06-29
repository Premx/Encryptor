using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Net.NetworkInformation;

namespace outlookaddin
{
    public partial class ThisAddIn
    {
        private static Settings settings;
        public static Outlook.Application AppObj;

        private static EncryptorForm MF = new EncryptorForm();


        public static void blankPassword()
        {
            settings.Password = "";
        }

        public static bool usePassword => settings.SavePassword;

        public static void setusepassword(bool b)
        {
            settings.SavePassword = b;
            settings.Save();
        }

        public static String getPassword()
        {
            String temp = settings.Password;
           return temp;
        }
        public static void setPassword(String pass)
        {
            settings.Password = pass;
            settings.Save();
        }
        
        private readonly String macaddress = GetMacAddress();



        private static string GetMacAddress()
        {
            const int MIN_MAC_ADDR_LENGTH = 12;
            string macAddress = string.Empty;
            long maxSpeed = -1;

            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                string tempMac = nic.GetPhysicalAddress().ToString();
                if (nic.Speed > maxSpeed &&
                    !string.IsNullOrEmpty(tempMac) &&
                    tempMac.Length >= MIN_MAC_ADDR_LENGTH)
                {
                   
                    maxSpeed = nic.Speed;
                    macAddress = tempMac;
                }
            }

            return macAddress;
        }

        public static void showMainForm()
        {
            MF.Close();
            MF = null;
            MF = new EncryptorForm();
            MF.Show();

        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AppObj = Application;

            settings = new Settings();
            
            
            this.Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(OutlookHelperTool.InspectorWrapper);

            //this.Application.ItemLoad += new Outlook.ApplicationEvents_11_ItemLoadEventHandler(OutlookHelperTool.ItemLoadEventHandler);

        }

        

            private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
