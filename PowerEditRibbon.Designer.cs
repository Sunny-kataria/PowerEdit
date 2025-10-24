namespace PowerEditAddIn
{
    partial class PowerEditRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PowerEditRibbon()
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
            this.PowerEdit = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Panel = this.Factory.CreateRibbonLabel();
            this.btnPowerEdit = this.Factory.CreateRibbonButton();
            this.PowerEdit.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // PowerEdit
            // 
            this.PowerEdit.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.PowerEdit.Groups.Add(this.group1);
            this.PowerEdit.Label = "POWEREDIT_TEST";
            this.PowerEdit.Name = "PowerEdit";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Panel);
            this.group1.Items.Add(this.btnPowerEdit);
            this.group1.Label = "Panel";
            this.group1.Name = "group1";
            // 
            // Panel
            // 
            this.Panel.Label = "label1";
            this.Panel.Name = "Panel";
            // 
            // btnPowerEdit
            // 
            this.btnPowerEdit.Label = "Open/Close Panel";
            this.btnPowerEdit.Name = "btnPowerEdit";
            this.btnPowerEdit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPowerEdit_Click);
            // 
            // PowerEditRibbon
            // 
            this.Name = "PowerEditRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.PowerEdit);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.PowerEdit.ResumeLayout(false);
            this.PowerEdit.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PowerEdit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel Panel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPowerEdit;
    }

    partial class ThisRibbonCollection
    {
        internal PowerEditRibbon Ribbon1
        {
            get { return this.GetRibbon<PowerEditRibbon>(); }
        }
    }
}
