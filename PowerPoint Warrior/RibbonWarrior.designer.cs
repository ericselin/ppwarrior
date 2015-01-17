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
            this.label1 = this.Factory.CreateRibbonLabel();
            this.editBoxGoToSlide = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.galleryStyles = this.Factory.CreateRibbonGallery();
            this.btnSaveStyle = this.Factory.CreateRibbonButton();
            this.btnDeleteStyle = this.Factory.CreateRibbonButton();
            this.btnPasteFromExcel = this.Factory.CreateRibbonButton();
            this.toggleAutoFit = this.Factory.CreateRibbonToggleButton();
            this.toggleWordWrap = this.Factory.CreateRibbonToggleButton();
            this.btnLineBelow = this.Factory.CreateRibbonButton();
            this.btnFormatBullets = this.Factory.CreateRibbonButton();
            this.btnHeaderLine = this.Factory.CreateRibbonButton();
            this.menuSetLanguage = this.Factory.CreateRibbonMenu();
            this.btnSetAlltext = this.Factory.CreateRibbonButton();
            this.btnSetLanguageEnglish = this.Factory.CreateRibbonButton();
            this.btnSetLanguageFinnsh = this.Factory.CreateRibbonButton();
            this.gallerySelectSimilar = this.Factory.CreateRibbonGallery();
            this.buttonSelectSimilarColorLine = this.Factory.CreateRibbonButton();
            this.buttonSelectSimilarColor = this.Factory.CreateRibbonButton();
            this.buttonSelectSimilarLine = this.Factory.CreateRibbonButton();
            this.buttonSelectSimilarWidth = this.Factory.CreateRibbonButton();
            this.buttonSelectSimilarHeight = this.Factory.CreateRibbonButton();
            this.buttonSelectSimilarHorizontal = this.Factory.CreateRibbonButton();
            this.buttonSelectSimilarVertical = this.Factory.CreateRibbonButton();
            this.galleryAlign = this.Factory.CreateRibbonGallery();
            this.buttonAlignTopToBottom = this.Factory.CreateRibbonButton();
            this.buttonAlignLeftToRight = this.Factory.CreateRibbonButton();
            this.btnSameHeight = this.Factory.CreateRibbonButton();
            this.btnSameWidth = this.Factory.CreateRibbonButton();
            this.btnSwapPos = this.Factory.CreateRibbonButton();
            this.btnPickUpPosition = this.Factory.CreateRibbonButton();
            this.btnApplyPosition = this.Factory.CreateRibbonButton();
            this.btnSplitShape = this.Factory.CreateRibbonButton();
            this.btnRemoveEffects = this.Factory.CreateRibbonButton();
            this.btnFormatTable = this.Factory.CreateRibbonButton();
            this.btnRemoveNotes = this.Factory.CreateRibbonButton();
            this.btnRemoveAnimations = this.Factory.CreateRibbonButton();
            this.btnFeedback = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnUpgrade = this.Factory.CreateRibbonButton();
            this.tabWarrior.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabWarrior
            // 
            this.tabWarrior.Groups.Add(this.group1);
            this.tabWarrior.Groups.Add(this.group2);
            this.tabWarrior.Groups.Add(this.group4);
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
            this.group1.Items.Add(this.btnFormatBullets);
            this.group1.Items.Add(this.btnLineBelow);
            this.group1.Items.Add(this.btnHeaderLine);
            this.group1.Items.Add(this.menuSetLanguage);
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.editBoxGoToSlide);
            this.group1.Items.Add(this.btnFormatTable);
            this.group1.Items.Add(this.gallerySelectSimilar);
            this.group1.Label = "Tools and Formatting";
            this.group1.Name = "group1";
            // 
            // label1
            // 
            this.label1.Label = "Go To Slide:";
            this.label1.Name = "label1";
            // 
            // editBoxGoToSlide
            // 
            this.editBoxGoToSlide.KeyTip = "GT";
            this.editBoxGoToSlide.Label = "Go To Slide";
            this.editBoxGoToSlide.Name = "editBoxGoToSlide";
            this.editBoxGoToSlide.OfficeImageId = "SlideTransitionApplyToAll";
            this.editBoxGoToSlide.ScreenTip = "Go to slide number";
            this.editBoxGoToSlide.ShowImage = true;
            this.editBoxGoToSlide.ShowLabel = false;
            this.editBoxGoToSlide.SuperTip = "Enter slide number of slide to show and press enter";
            this.editBoxGoToSlide.Text = null;
            this.editBoxGoToSlide.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBoxGoToSlide_TextChanged);
            // 
            // group2
            // 
            this.group2.Items.Add(this.galleryAlign);
            this.group2.Items.Add(this.btnSameHeight);
            this.group2.Items.Add(this.btnSameWidth);
            this.group2.Items.Add(this.btnSwapPos);
            this.group2.Items.Add(this.btnPickUpPosition);
            this.group2.Items.Add(this.btnApplyPosition);
            this.group2.Items.Add(this.btnSplitShape);
            this.group2.Label = "Size and Position";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnFeedback);
            this.group3.Items.Add(this.btnAbout);
            this.group3.Items.Add(this.btnUpgrade);
            this.group3.Name = "group3";
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnRemoveEffects);
            this.group4.Items.Add(this.btnRemoveNotes);
            this.group4.Items.Add(this.btnRemoveAnimations);
            this.group4.Label = "Cleanup";
            this.group4.Name = "group4";
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
            // btnLineBelow
            // 
            this.btnLineBelow.Image = global::PowerPoint_Warrior.Properties.Resources.IconLineBelow;
            this.btnLineBelow.KeyTip = "LB";
            this.btnLineBelow.Label = "Line Below";
            this.btnLineBelow.Name = "btnLineBelow";
            this.btnLineBelow.ScreenTip = "Insert a line below selected shape(s)";
            this.btnLineBelow.ShowImage = true;
            this.btnLineBelow.SuperTip = "Note that the line used will be the default line format for the current presentat" +
    "ion";
            this.btnLineBelow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLineBelow_Click);
            // 
            // btnFormatBullets
            // 
            this.btnFormatBullets.Image = global::PowerPoint_Warrior.Properties.Resources.IconFormatBulletList;
            this.btnFormatBullets.KeyTip = "B";
            this.btnFormatBullets.Label = "Format Bullets";
            this.btnFormatBullets.Name = "btnFormatBullets";
            this.btnFormatBullets.ScreenTip = "Format bullets to look professional";
            this.btnFormatBullets.ShowImage = true;
            this.btnFormatBullets.SuperTip = "Formatting can be done on several shapes at once, for one shape, or for selected " +
    "text only";
            this.btnFormatBullets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatBullets_Click);
            // 
            // btnHeaderLine
            // 
            this.btnHeaderLine.Image = global::PowerPoint_Warrior.Properties.Resources.IconHeaderLine;
            this.btnHeaderLine.KeyTip = "LL";
            this.btnHeaderLine.Label = "Header Line";
            this.btnHeaderLine.Name = "btnHeaderLine";
            this.btnHeaderLine.ScreenTip = "Create / align header line";
            this.btnHeaderLine.ShowImage = true;
            this.btnHeaderLine.SuperTip = "Creates a trailing header line to a text box, or if two objects are selected alig" +
    "ns the header line according to the object below";
            this.btnHeaderLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHeaderLine_Click);
            // 
            // menuSetLanguage
            // 
            this.menuSetLanguage.Items.Add(this.btnSetAlltext);
            this.menuSetLanguage.Items.Add(this.btnSetLanguageEnglish);
            this.menuSetLanguage.Items.Add(this.btnSetLanguageFinnsh);
            this.menuSetLanguage.KeyTip = "U";
            this.menuSetLanguage.Label = "Language";
            this.menuSetLanguage.Name = "menuSetLanguage";
            this.menuSetLanguage.OfficeImageId = "SetLanguage";
            this.menuSetLanguage.ScreenTip = "Set language of entire presentation";
            this.menuSetLanguage.ShowImage = true;
            this.menuSetLanguage.SuperTip = "Sets the language of the entire presentation, including grouped objects";
            // 
            // btnSetAlltext
            // 
            this.btnSetAlltext.Enabled = false;
            this.btnSetAlltext.Label = "Set whole presentation to:";
            this.btnSetAlltext.Name = "btnSetAlltext";
            this.btnSetAlltext.ShowImage = true;
            // 
            // btnSetLanguageEnglish
            // 
            this.btnSetLanguageEnglish.KeyTip = "E";
            this.btnSetLanguageEnglish.Label = "English";
            this.btnSetLanguageEnglish.Name = "btnSetLanguageEnglish";
            this.btnSetLanguageEnglish.ShowImage = true;
            this.btnSetLanguageEnglish.Tag = "";
            this.btnSetLanguageEnglish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetLanguage_Click);
            // 
            // btnSetLanguageFinnsh
            // 
            this.btnSetLanguageFinnsh.KeyTip = "F";
            this.btnSetLanguageFinnsh.Label = "Finnish";
            this.btnSetLanguageFinnsh.Name = "btnSetLanguageFinnsh";
            this.btnSetLanguageFinnsh.ShowImage = true;
            this.btnSetLanguageFinnsh.Tag = "";
            this.btnSetLanguageFinnsh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetLanguage_Click);
            // 
            // gallerySelectSimilar
            // 
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarColorLine);
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarColor);
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarLine);
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarWidth);
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarHeight);
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarHorizontal);
            this.gallerySelectSimilar.Buttons.Add(this.buttonSelectSimilarVertical);
            this.gallerySelectSimilar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallerySelectSimilar.KeyTip = "SS";
            this.gallerySelectSimilar.Label = "Select similar";
            this.gallerySelectSimilar.Name = "gallerySelectSimilar";
            this.gallerySelectSimilar.OfficeImageId = "SelectionPane";
            this.gallerySelectSimilar.ScreenTip = "Select similar shapes";
            this.gallerySelectSimilar.ShowImage = true;
            this.gallerySelectSimilar.SuperTip = "Select all shapes on the slide which match certain features of the currently sele" +
    "cted shape";
            this.gallerySelectSimilar.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallerySelectSimilar_ButtonClick);
            // 
            // buttonSelectSimilarColorLine
            // 
            this.buttonSelectSimilarColorLine.KeyTip = "E";
            this.buttonSelectSimilarColorLine.Label = "Same Color and Line";
            this.buttonSelectSimilarColorLine.Name = "buttonSelectSimilarColorLine";
            // 
            // buttonSelectSimilarColor
            // 
            this.buttonSelectSimilarColor.Label = "Same Color";
            this.buttonSelectSimilarColor.Name = "buttonSelectSimilarColor";
            // 
            // buttonSelectSimilarLine
            // 
            this.buttonSelectSimilarLine.Label = "Same Line";
            this.buttonSelectSimilarLine.Name = "buttonSelectSimilarLine";
            // 
            // buttonSelectSimilarWidth
            // 
            this.buttonSelectSimilarWidth.Label = "Same Width (+-10%)";
            this.buttonSelectSimilarWidth.Name = "buttonSelectSimilarWidth";
            // 
            // buttonSelectSimilarHeight
            // 
            this.buttonSelectSimilarHeight.Label = "Same Height (+-10%)";
            this.buttonSelectSimilarHeight.Name = "buttonSelectSimilarHeight";
            // 
            // buttonSelectSimilarHorizontal
            // 
            this.buttonSelectSimilarHorizontal.Label = "On Same Horizontal (+-15px)";
            this.buttonSelectSimilarHorizontal.Name = "buttonSelectSimilarHorizontal";
            // 
            // buttonSelectSimilarVertical
            // 
            this.buttonSelectSimilarVertical.Label = "On Same Vertical (+-15px)";
            this.buttonSelectSimilarVertical.Name = "buttonSelectSimilarVertical";
            // 
            // galleryAlign
            // 
            this.galleryAlign.Buttons.Add(this.buttonAlignTopToBottom);
            this.galleryAlign.Buttons.Add(this.buttonAlignLeftToRight);
            this.galleryAlign.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.galleryAlign.KeyTip = "GA";
            this.galleryAlign.Label = "Align";
            this.galleryAlign.Name = "galleryAlign";
            this.galleryAlign.OfficeImageId = "ObjectsUngroup";
            this.galleryAlign.ScreenTip = "Align objects";
            this.galleryAlign.ShowImage = true;
            this.galleryAlign.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryAlign_ButtonClick);
            // 
            // buttonAlignTopToBottom
            // 
            this.buttonAlignTopToBottom.KeyTip = "T";
            this.buttonAlignTopToBottom.Label = "Top to bottom";
            this.buttonAlignTopToBottom.Name = "buttonAlignTopToBottom";
            this.buttonAlignTopToBottom.OfficeImageId = "ObjectsAlignTop";
            this.buttonAlignTopToBottom.ScreenTip = "Align objects top to bottom";
            this.buttonAlignTopToBottom.ShowImage = true;
            this.buttonAlignTopToBottom.SuperTip = "Align top of objects to bottom of topmost object";
            // 
            // buttonAlignLeftToRight
            // 
            this.buttonAlignLeftToRight.KeyTip = "L";
            this.buttonAlignLeftToRight.Label = "Left to right";
            this.buttonAlignLeftToRight.Name = "buttonAlignLeftToRight";
            this.buttonAlignLeftToRight.OfficeImageId = "ObjectsAlignLeft";
            this.buttonAlignLeftToRight.ScreenTip = "Align objects left to right";
            this.buttonAlignLeftToRight.ShowImage = true;
            this.buttonAlignLeftToRight.SuperTip = "Align left edge of objects to right of leftmost object";
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
            // btnPickUpPosition
            // 
            this.btnPickUpPosition.KeyTip = "P";
            this.btnPickUpPosition.Label = "Pick Up Pos.";
            this.btnPickUpPosition.Name = "btnPickUpPosition";
            this.btnPickUpPosition.OfficeImageId = "PickUpStyle";
            this.btnPickUpPosition.ScreenTip = "Pick up position and size of selected shape";
            this.btnPickUpPosition.ShowImage = true;
            this.btnPickUpPosition.SuperTip = "Picks up position and size, which can then be used to set another shape to the sa" +
    "me position (possibly on another slide) using \"Apply Pos.\"";
            this.btnPickUpPosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPickUpPosition_Click);
            // 
            // btnApplyPosition
            // 
            this.btnApplyPosition.Enabled = false;
            this.btnApplyPosition.KeyTip = "A";
            this.btnApplyPosition.Label = "Apply Pos.";
            this.btnApplyPosition.Name = "btnApplyPosition";
            this.btnApplyPosition.OfficeImageId = "PasteApplyStyle";
            this.btnApplyPosition.ScreenTip = "Move selected shape to the picked up position and set same width and height";
            this.btnApplyPosition.ShowImage = true;
            this.btnApplyPosition.SuperTip = "Moves the upper-right corner of the selected shape to the same place and set the " +
    "same width and height as the shape for which the position and size was picked up" +
    " using \"Pick Up Pos.\"";
            this.btnApplyPosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyPosition_Click);
            // 
            // btnSplitShape
            // 
            this.btnSplitShape.KeyTip = "SP";
            this.btnSplitShape.Label = "Split shape";
            this.btnSplitShape.Name = "btnSplitShape";
            this.btnSplitShape.OfficeImageId = "TraceDependents";
            this.btnSplitShape.ScreenTip = "Split shape so that every paragraph becomes its own shape";
            this.btnSplitShape.ShowImage = true;
            this.btnSplitShape.SuperTip = "Split selected shape into equally sized shapes with one paragraph per shape";
            this.btnSplitShape.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSplitShape_Click);
            // 
            // btnRemoveEffects
            // 
            this.btnRemoveEffects.KeyTip = "RE";
            this.btnRemoveEffects.Label = "Remove effects";
            this.btnRemoveEffects.Name = "btnRemoveEffects";
            this.btnRemoveEffects.OfficeImageId = "FormatPainter";
            this.btnRemoveEffects.ScreenTip = "Remove all effects applied to the selected shape(s)";
            this.btnRemoveEffects.ShowImage = true;
            this.btnRemoveEffects.SuperTip = "Removes shadows, 3D rotations, and all other visual effects";
            this.btnRemoveEffects.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveEffects_Click);
            // 
            // btnFormatTable
            // 
            this.btnFormatTable.KeyTip = "T";
            this.btnFormatTable.Label = "Format Table";
            this.btnFormatTable.Name = "btnFormatTable";
            this.btnFormatTable.OfficeImageId = "FormatAsTableGallery";
            this.btnFormatTable.ScreenTip = "Format table to basic style";
            this.btnFormatTable.ShowImage = true;
            this.btnFormatTable.SuperTip = "Removes background colors and sets grid to thin lines";
            this.btnFormatTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatTable_Click);
            // 
            // btnRemoveNotes
            // 
            this.btnRemoveNotes.Label = "Remove notes";
            this.btnRemoveNotes.Name = "btnRemoveNotes";
            this.btnRemoveNotes.OfficeImageId = "ReviewDeleteAllMarkupOnSlide";
            this.btnRemoveNotes.ScreenTip = "Remove notes from all slides";
            this.btnRemoveNotes.ShowImage = true;
            this.btnRemoveNotes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveNotes_Click);
            // 
            // btnRemoveAnimations
            // 
            this.btnRemoveAnimations.Label = "Remove animations";
            this.btnRemoveAnimations.Name = "btnRemoveAnimations";
            this.btnRemoveAnimations.OfficeImageId = "CDAudioStopTime";
            this.btnRemoveAnimations.ScreenTip = "Remove animations from all slides";
            this.btnRemoveAnimations.ShowImage = true;
            this.btnRemoveAnimations.SuperTip = "Removes slide transitions as well as all shape animations";
            this.btnRemoveAnimations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveAnimations_Click);
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
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveEffects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLineBelow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatBullets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuSetLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetAlltext;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetLanguageEnglish;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetLanguageFinnsh;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxGoToSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallerySelectSimilar;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarColorLine;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarColor;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarLine;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarWidth;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarHeight;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarHorizontal;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonSelectSimilarVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryAlign;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonAlignTopToBottom;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonAlignLeftToRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPickUpPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplyPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSplitShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHeaderLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveNotes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveAnimations;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonWarrior RibbonWarrior
        {
            get { return this.GetRibbon<RibbonWarrior>(); }
        }
    }
}
