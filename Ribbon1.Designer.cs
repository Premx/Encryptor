namespace outlookaddin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.praktikumaddin = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonCrypt = this.Factory.CreateRibbonButton();
            this.checkBoxSavePassword = this.Factory.CreateRibbonCheckBox();
            this.praktikumaddin.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // praktikumaddin
            // 
            this.praktikumaddin.Groups.Add(this.group1);
            this.praktikumaddin.Label = "Encryptor";
            this.praktikumaddin.Name = "praktikumaddin";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonCrypt);
            this.group1.Items.Add(this.checkBoxSavePassword);
            this.group1.Name = "group1";
            // 
            // buttonCrypt
            // 
            this.buttonCrypt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonCrypt.Image = global::outlookaddin.Properties.Resources._1498570529_icon_117_lock_stripes;
            this.buttonCrypt.Label = "Encrytpor";
            this.buttonCrypt.Name = "buttonCrypt";
            this.buttonCrypt.ShowImage = true;
            this.buttonCrypt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // checkBoxSavePassword
            // 
            this.checkBoxSavePassword.Label = "Save Password";
            this.checkBoxSavePassword.Name = "checkBoxSavePassword";
            this.checkBoxSavePassword.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxSavePassword_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" +
    "l.Read";
            this.Tabs.Add(this.praktikumaddin);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.praktikumaddin.ResumeLayout(false);
            this.praktikumaddin.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab praktikumaddin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCrypt;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSavePassword;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
