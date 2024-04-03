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
        private SlideService _slideClass;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            //
            var container = ContainerConfig.RegisterServices();
            _slideClass = container.Resolve<SlideService>();
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
            this.groupText = this.Factory.CreateRibbonGroup();
            this.fSizePlus = this.Factory.CreateRibbonButton();
            this.fSizeMinus = this.Factory.CreateRibbonButton();
            this.copyTxt = this.Factory.CreateRibbonButton();
            this.pasteTxt = this.Factory.CreateRibbonButton();
            this.groupShape = this.Factory.CreateRibbonGroup();
            this.copyPosition = this.Factory.CreateRibbonButton();
            this.pastePosition = this.Factory.CreateRibbonButton();
            this.alignLeft = this.Factory.CreateRibbonButton();
            this.alignTop = this.Factory.CreateRibbonButton();
            this.alignRight = this.Factory.CreateRibbonButton();
            this.alignBottom = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.groupText.SuspendLayout();
            this.groupShape.SuspendLayout();
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
            this.tab2.Groups.Add(this.groupText);
            this.tab2.Groups.Add(this.groupShape);
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
            // groupText
            // 
            this.groupText.Items.Add(this.fSizePlus);
            this.groupText.Items.Add(this.fSizeMinus);
            this.groupText.Items.Add(this.copyTxt);
            this.groupText.Items.Add(this.pasteTxt);
            this.groupText.Label = "Text ";
            this.groupText.Name = "groupText";
            // 
            // fSizePlus
            // 
            this.fSizePlus.Label = "Font Size +";
            this.fSizePlus.Name = "fSizePlus";
            this.fSizePlus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fSizePlus_Click);
            // 
            // fSizeMinus
            // 
            this.fSizeMinus.Label = "Font Size -";
            this.fSizeMinus.Name = "fSizeMinus";
            this.fSizeMinus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fSizeMinus_Click);
            // 
            // copyTxt
            // 
            this.copyTxt.Label = "Copy Text";
            this.copyTxt.Name = "copyTxt";
            this.copyTxt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyTxt_Click);
            // 
            // pasteTxt
            // 
            this.pasteTxt.Label = "Paste Text";
            this.pasteTxt.Name = "pasteTxt";
            this.pasteTxt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteTxt_Click);
            // 
            // groupShape
            // 
            this.groupShape.Items.Add(this.copyPosition);
            this.groupShape.Items.Add(this.pastePosition);
            this.groupShape.Items.Add(this.alignLeft);
            this.groupShape.Items.Add(this.alignTop);
            this.groupShape.Items.Add(this.alignRight);
            this.groupShape.Items.Add(this.alignBottom);
            this.groupShape.Label = "Shape";
            this.groupShape.Name = "groupShape";
            // 
            // copyPosition
            // 
            this.copyPosition.Label = "CopyPosition";
            this.copyPosition.Name = "copyPosition";
            this.copyPosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyPosition_Click);
            // 
            // pastePosition
            // 
            this.pastePosition.Label = "PastePosition";
            this.pastePosition.Name = "pastePosition";
            this.pastePosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pastePosition_Click);
            // 
            // alignLeft
            // 
            this.alignLeft.Label = "Align Left";
            this.alignLeft.Name = "alignLeft";
            this.alignLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignLeft_Click);
            // 
            // alignTop
            // 
            this.alignTop.Label = "Align Top";
            this.alignTop.Name = "alignTop";
            this.alignTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignTop_Click);
            // 
            // alignRight
            // 
            this.alignRight.Label = "Align Right";
            this.alignRight.Name = "alignRight";
            this.alignRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignRight_Click);
            // 
            // alignBottom
            // 
            this.alignBottom.Label = "Align Bottom";
            this.alignBottom.Name = "alignBottom";
            this.alignBottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignBottom_Click);
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
            this.groupText.ResumeLayout(false);
            this.groupText.PerformLayout();
            this.groupShape.ResumeLayout(false);
            this.groupShape.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fSizePlus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fSizeMinus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupShape;
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
