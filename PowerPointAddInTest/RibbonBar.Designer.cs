namespace PowerPointAddInTest
{
    partial class RibbonBar : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonBar()
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
            this.tabChangeStyle = this.Factory.CreateRibbonTab();
            this.groupStart = this.Factory.CreateRibbonGroup();
            this.buttonStart = this.Factory.CreateRibbonButton();
            this.tabChangeStyle.SuspendLayout();
            this.groupStart.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabChangeStyle
            // 
            this.tabChangeStyle.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabChangeStyle.Groups.Add(this.groupStart);
            this.tabChangeStyle.Label = "Change style addin";
            this.tabChangeStyle.Name = "tabChangeStyle";
            // 
            // groupStart
            // 
            this.groupStart.Items.Add(this.buttonStart);
            this.groupStart.Label = "Change";
            this.groupStart.Name = "groupStart";
            // 
            // buttonStart
            // 
            this.buttonStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStart.Image = global::PowerPointAddInTest.Properties.Resources.StartImage;
            this.buttonStart.Label = "Start";
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.ShowImage = true;
            this.buttonStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStart_Click);
            // 
            // RibbonBar
            // 
            this.Name = "RibbonBar";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabChangeStyle);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonBar_Load);
            this.tabChangeStyle.ResumeLayout(false);
            this.tabChangeStyle.PerformLayout();
            this.groupStart.ResumeLayout(false);
            this.groupStart.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabChangeStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStart;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonBar RibbonBar
        {
            get { return this.GetRibbon<RibbonBar>(); }
        }
    }
}
