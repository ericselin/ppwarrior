namespace PowerPoint_Warrior
{
    partial class RibbonWarrior : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonWarrior()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            this.tabWarrior = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.galleryStyles = this.Factory.CreateRibbonGallery();
            this.btnSaveStyle = this.Factory.CreateRibbonButton();
            this.btnDeleteStyle = this.Factory.CreateRibbonButton();
            this.btnPasteFromExcel = this.Factory.CreateRibbonButton();
            this.toggleAutoFit = this.Factory.CreateRibbonToggleButton();
            this.toggleWordWrap = this.Factory.CreateRibbonToggleButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.gallery1 = this.Factory.CreateRibbonGallery();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.gallery2 = this.Factory.CreateRibbonGallery();
            this.btnSameHeight = this.Factory.CreateRibbonButton();
            this.btnSameWidth = this.Factory.CreateRibbonButton();
            this.btnSwapPos = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnFeedback = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnUpgrade = this.Factory.CreateRibbonButton();
            this.tabWarrior.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabWarrior
            // 
            this.tabWarrior.Groups.Add(this.group1);
            this.tabWarrior.Groups.Add(this.group2);
            this.tabWarrior.Groups.Add(this.group3);
            this.tabWarrior.KeyTip = "C";
            this.tabWarrior.Label = "Warrior";
            this.tabWarrior.Name = "tabWarrior";
            // 
            // group1
            // 
            this.group1.Items.Add(this.galleryStyles);
            this.group1.Items.Add(this.btnPasteFromExcel);
            this.group1.Items.Add(this.toggleAutoFit);
            this.group1.Items.Add(this.toggleWordWrap);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.gallery1);
            this.group1.Label = "Tools and Formatting";
            this.group1.Name = "group1";
            // 
            // galleryStyles
            // 
            this.galleryStyles.Buttons.Add(this.btnSaveStyle);
            this.galleryStyles.Buttons.Add(this.btnDeleteStyle);
            this.galleryStyles.ColumnCount = 1;
            this.galleryStyles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            ribbonDropDownItemImpl1.Label = "(no styles)";
            this.galleryStyles.Items.Add(ribbonDropDownItemImpl1);
            this.galleryStyles.KeyTip = "S";
            this.galleryStyles.Label = "Styles";
            this.galleryStyles.Name = "galleryStyles";
            this.galleryStyles.OfficeImageId = "ShapeStylesGallery";
            this.galleryStyles.ShowImage = true;
            this.galleryStyles.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryStyles_ButtonClick);
            this.galleryStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryStyles_Click);
            // 
            // btnSaveStyle
            // 
            this.btnSaveStyle.Label = "Save current style";
            this.btnSaveStyle.Name = "btnSaveStyle";
            this.btnSaveStyle.OfficeImageId = "FileSave";
            this.btnSaveStyle.ScreenTip = "Save style of selected shape";
            this.btnSaveStyle.ShowImage = true;
            this.btnSaveStyle.SuperTip = "You can both save as a new style or overwrite an existing style";
            // 
            // btnDeleteStyle
            // 
            this.btnDeleteStyle.Label = "Delete style";
            this.btnDeleteStyle.Name = "btnDeleteStyle";
            this.btnDeleteStyle.OfficeImageId = "Delete";
            this.btnDeleteStyle.ShowImage = true;
            // 
            // btnPasteFromExcel
            // 
            this.btnPasteFromExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPasteFromExcel.KeyTip = "E";
            this.btnPasteFromExcel.Label = "Paste from Excel";
            this.btnPasteFromExcel.Name = "btnPasteFromExcel";
            this.btnPasteFromExcel.OfficeImageId = "ImportExcel";
            this.btnPasteFromExcel.ScreenTip = "Paste data from Excel to a PowerPoint table";
            this.btnPasteFromExcel.ShowImage = true;
            this.btnPasteFromExcel.SuperTip = "Paste data from Excel while keeping the formatting of the PowerPoint table";
            this.btnPasteFromExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPasteFromExcel_Click);
            // 
            // toggleAutoFit
            // 
            this.toggleAutoFit.Image = global::PowerPoint_Warrior.Properties.Resources.IconResize;
            this.toggleAutoFit.KeyTip = "A";
            this.toggleAutoFit.Label = "Toggle AutoFit";
            this.toggleAutoFit.Name = "toggleAutoFit";
            this.toggleAutoFit.ScreenTip = "Toggle shape AutoFit from \"do not AutoFit\" to \"resize shape to fit text\"";
            this.toggleAutoFit.ShowImage = true;
            this.toggleAutoFit.SuperTip = "If selected shape(s) are set to \"shrink text\" or have different AutoFit settings," +
    " the icon will be grey (but the button still works)";
            this.toggleAutoFit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleAutoFit_Click);
            // 
            // toggleWordWrap
            // 
            this.toggleWordWrap.Image = global::PowerPoint_Warrior.Properties.Resources.IconWordwrap;
            this.toggleWordWrap.KeyTip = "V";
            this.toggleWordWrap.Label = "Toggle Word Wrap";
            this.toggleWordWrap.Name = "toggleWordWrap";
            this.toggleWordWrap.ScreenTip = "Toggle word wrapping in selected shapes";
            this.toggleWordWrap.ShowImage = true;
            this.toggleWordWrap.SuperTip = "If selected shapes have different word wrapping properties, the icon will be grey" +
    " (but the button will still work)";
            this.toggleWordWrap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleWordWrap_Click);
            // 
            // button1
            // 
            this.button1.Label = "Remove effects";
            this.button1.Name = "button1";
            this.button1.Visible = false;
            // 
            // button2
            // 
            this.button2.Label = "Line Below";
            this.button2.Name = "button2";
            this.button2.Visible = false;
            // 
            // button3
            // 
            this.button3.Label = "Format Bullets";
            this.button3.Name = "button3";
            this.button3.Visible = false;
            // 
            // button4
            // 
            this.button4.Label = "Format Table";
            this.button4.Name = "button4";
            this.button4.Visible = false;
            // 
            // gallery1
            // 
            this.gallery1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery1.Label = "Select Similar";
            this.gallery1.Name = "gallery1";
            this.gallery1.ShowImage = true;
            this.gallery1.Visible = false;
            // 
            // group2
            // 
            this.group2.Items.Add(this.gallery2);
            this.group2.Items.Add(this.btnSameHeight);
            this.group2.Items.Add(this.btnSameWidth);
            this.group2.Items.Add(this.btnSwapPos);
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button6);
            this.group2.Items.Add(this.button7);
            this.group2.Label = "Size and Position";
            this.group2.Name = "group2";
            // 
            // gallery2
            // 
            this.gallery2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery2.Label = "Align";
            this.gallery2.Name = "gallery2";
            this.gallery2.ShowImage = true;
            this.gallery2.Visible = false;
            // 
            // btnSameHeight
            // 
            this.btnSameHeight.KeyTip = "H";
            this.btnSameHeight.Label = "Same Height";
            this.btnSameHeight.Name = "btnSameHeight";
            this.btnSameHeight.OfficeImageId = "SizeToControlHeight";
            this.btnSameHeight.ScreenTip = "Set shapes to same height";
            this.btnSameHeight.ShowImage = true;
            this.btnSameHeight.SuperTip = "Sets all the selected shapes to the same height as the shape that was selected fi" +
    "rst";
            this.btnSameHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSameHeightOrWidth_Click);
            // 
            // btnSameWidth
            // 
            this.btnSameWidth.KeyTip = "W";
            this.btnSameWidth.Label = "Same Width";
            this.btnSameWidth.Name = "btnSameWidth";
            this.btnSameWidth.OfficeImageId = "SizeToControlWidth";
            this.btnSameWidth.ScreenTip = "Set shapes to same width";
            this.btnSameWidth.ShowImage = true;
            this.btnSameWidth.SuperTip = "Sets all the selected shapes to the same width as the shape that was selected fir" +
    "st";
            this.btnSameWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSameHeightOrWidth_Click);
            // 
            // btnSwapPos
            // 
            this.btnSwapPos.KeyTip = "P";
            this.btnSwapPos.Label = "Swap Positions";
            this.btnSwapPos.Name = "btnSwapPos";
            this.btnSwapPos.OfficeImageId = "CircularReferences";
            this.btnSwapPos.ScreenTip = "Swap positions of selected shapes";
            this.btnSwapPos.ShowImage = true;
            this.btnSwapPos.SuperTip = "Swaps the positions (upper-left corner) of two shapes";
            this.btnSwapPos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSwapPos_Click);
            // 
            // button5
            // 
            this.button5.Label = "Pick Up Position";
            this.button5.Name = "button5";
            this.button5.Visible = false;
            // 
            // button6
            // 
            this.button6.Label = "Apply Position";
            this.button6.Name = "button6";
            this.button6.Visible = false;
            // 
            // button7
            // 
            this.button7.Label = "Split Shape";
            this.button7.Name = "button7";
            this.button7.Visible = false;
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnFeedback);
            this.group3.Items.Add(this.btnAbout);
            this.group3.Items.Add(this.btnUpgrade);
            this.group3.Name = "group3";
            // 
            // btnFeedback
            // 
            this.btnFeedback.Label = "Send Feedback";
            this.btnFeedback.Name = "btnFeedback";
            this.btnFeedback.OfficeImageId = "PostReplyToFolder";
            this.btnFeedback.ScreenTip = "Send feedback to the developers";
            this.btnFeedback.ShowImage = true;
            this.btnFeedback.SuperTip = "Thank you for your feedback - whatever it may be!";
            this.btnFeedback.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFeedback_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "Settings and About";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.OfficeImageId = "Info";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // btnUpgrade
            // 
            this.btnUpgrade.Label = "Upgrade";
            this.btnUpgrade.Name = "btnUpgrade";
            this.btnUpgrade.OfficeImageId = "ViewOnlineConnection";
            this.btnUpgrade.ScreenTip = "Upgrade product for more functionality";
            this.btnUpgrade.ShowImage = true;
            this.btnUpgrade.SuperTip = "If you wish to upgrade to a paid version, you can also send an e-mail to eric.sel" +
    "in@gmail.com";
            this.btnUpgrade.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpgrade_Click);
            // 
            // RibbonWarrior
            // 
            this.Name = "RibbonWarrior";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabWarrior);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonWarrior_Load);
            this.tabWarrior.ResumeLayout(false);
            this.tabWarrior.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabWarrior;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPasteFromExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSameHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSameWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSwapPos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFeedback;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleAutoFit;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleWordWrap;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryStyles;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveStyle;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpgrade;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonWarrior RibbonWarrior
        {
            get { return this.GetRibbon<RibbonWarrior>(); }
        }
    }
}
