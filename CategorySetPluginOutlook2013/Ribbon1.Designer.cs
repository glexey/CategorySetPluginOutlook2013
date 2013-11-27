namespace CategorySetPluginOutlook2013
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SetCategory = this.Factory.CreateRibbonGallery();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.SetCategory);
            this.group1.Label = "myPlugins";
            this.group1.Name = "group1";
            // 
            // SetCategory
            // 
            this.SetCategory.Buttons.Add(this.button1);
            this.SetCategory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SetCategory.Label = "Categorize";
            this.SetCategory.Name = "SetCategory";
            this.SetCategory.OfficeImageId = "ColorMenu";
            this.SetCategory.ShowImage = true;
            this.SetCategory.SuperTip = "Assign message(s) to a category(ies)";
            this.SetCategory.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery1_ButtonClick);
            this.SetCategory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery1_Click);
            this.SetCategory.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery1_ItemLoad);
            // 
            // button1
            // 
            this.button1.Label = "&All Categories...";
            this.button1.Name = "button1";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery SetCategory;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
