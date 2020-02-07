namespace WordAddIn_Segment
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_Segment = this.Factory.CreateRibbonButton();
            this.btn_Statistics = this.Factory.CreateRibbonButton();
            this.btn_About = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "Segmentation";
            this.tab2.Name = "tab2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_Segment);
            this.group1.Items.Add(this.btn_Statistics);
            this.group1.Label = "Main";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_About);
            this.group2.Label = "Help";
            this.group2.Name = "group2";
            // 
            // btn_Segment
            // 
            this.btn_Segment.Label = "Segment";
            this.btn_Segment.Name = "btn_Segment";
            this.btn_Segment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Segment_Click);
            // 
            // btn_Statistics
            // 
            this.btn_Statistics.Label = "Statistics";
            this.btn_Statistics.Name = "btn_Statistics";
            this.btn_Statistics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Statistics_Click);
            // 
            // btn_About
            // 
            this.btn_About.Label = "about";
            this.btn_About.Name = "btn_About";
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Segment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Statistics;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_About;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
