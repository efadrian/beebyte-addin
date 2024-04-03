using Unity;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private Application _pptApp;
        private SlideService _myClass;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            //
            var container = ContainerConfig.RegisterServices();
            _myClass = container.Resolve<SlideService>();
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
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.fSizePlus = this.Factory.CreateRibbonButton();
            this.fSizeMinus = this.Factory.CreateRibbonButton();
            this.copyTxt = this.Factory.CreateRibbonButton();
            this.pasteTxt = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.copyPosition = this.Factory.CreateRibbonButton();
            this.pastePosition = this.Factory.CreateRibbonButton();
            this.alignLeft = this.Factory.CreateRibbonButton();
            this.alignTop = this.Factory.CreateRibbonButton();
            this.alignRight = this.Factory.CreateRibbonButton();
            this.alignBottom = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Groups.Add(this.group4);
            this.tab2.Label = "BeeByte";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.button2);
            this.group2.Label = "Slide";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.Label = "Add";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Label = "Remove";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.fSizePlus);
            this.group3.Items.Add(this.fSizeMinus);
            this.group3.Items.Add(this.copyTxt);
            this.group3.Items.Add(this.pasteTxt);
            this.group3.Label = "Text ";
            this.group3.Name = "group3";
            // 
            // fSizePlus
            // 
            this.fSizePlus.Label = "Font Size +";
            this.fSizePlus.Name = "fSizePlus";
            // 
            // fSizeMinus
            // 
            this.fSizeMinus.Label = "Font Size -";
            this.fSizeMinus.Name = "fSizeMinus";
            // 
            // copyTxt
            // 
            this.copyTxt.Label = "Copy Text";
            this.copyTxt.Name = "copyTxt";
            // 
            // pasteTxt
            // 
            this.pasteTxt.Label = "Paste Text";
            this.pasteTxt.Name = "pasteTxt";
            // 
            // group4
            // 
            this.group4.Items.Add(this.copyPosition);
            this.group4.Items.Add(this.pastePosition);
            this.group4.Items.Add(this.alignLeft);
            this.group4.Items.Add(this.alignTop);
            this.group4.Items.Add(this.alignRight);
            this.group4.Items.Add(this.alignBottom);
            this.group4.Label = "Shape";
            this.group4.Name = "group4";
            // 
            // copyPosition
            // 
            this.copyPosition.Label = "CopyPosition";
            this.copyPosition.Name = "copyPosition";
            // 
            // pastePosition
            // 
            this.pastePosition.Label = "PastePosition";
            this.pastePosition.Name = "pastePosition";
            // 
            // alignLeft
            // 
            this.alignLeft.Label = "Align Left";
            this.alignLeft.Name = "alignLeft";
            // 
            // alignTop
            // 
            this.alignTop.Label = "Align Top";
            this.alignTop.Name = "alignTop";
            // 
            // alignRight
            // 
            this.alignRight.Label = "Align Right";
            this.alignRight.Name = "alignRight";
            // 
            // alignBottom
            // 
            this.alignBottom.Label = "Align Bottom";
            this.alignBottom.Name = "alignBottom";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fSizePlus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fSizeMinus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pastePosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignBottom;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
