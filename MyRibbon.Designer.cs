namespace UserStorySimilarityAddIn
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.tabUserStory = this.Factory.CreateRibbonTab();
            this.groupActions = this.Factory.CreateRibbonGroup();
            this.btnCompare = this.Factory.CreateRibbonButton();
            this.tabUserStory.SuspendLayout();
            this.groupActions.SuspendLayout();
            // 
            // tabUserStory
            // 
            this.tabUserStory.Label = "User Story Tools";
            this.tabUserStory.Name = "tabUserStory";
            this.tabUserStory.Groups.Add(this.groupActions);
            // 
            // groupActions
            // 
            this.groupActions.Label = "Actions";
            this.groupActions.Name = "groupActions";
            this.groupActions.Items.Add(this.btnCompare);
            // 
            // btnCompare
            // 
            this.btnCompare.Label = "Compare Stories";
            this.btnCompare.Name = "btnCompare";
            this.btnCompare.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareButton_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabUserStory);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tabUserStory.ResumeLayout(false);
            this.tabUserStory.PerformLayout();
            this.groupActions.ResumeLayout(false);
            this.groupActions.PerformLayout();
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabUserStory;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompare;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
