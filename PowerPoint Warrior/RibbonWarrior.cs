using System;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Warrior_Common;
using Style_Manager;

namespace PowerPoint_Warrior
{
    public partial class RibbonWarrior
    {
        Style_Manager.StyleLogic styles;
        string officeVersion;
        string userEmail;
        UsageLogger logger;
        bool licenseValid;

        private void RibbonWarrior_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                // get version and if 2013, make tab CAPS
                officeVersion = Globals.ThisAddIn.Application.Version;
                if (float.Parse(officeVersion, System.Globalization.CultureInfo.InvariantCulture) >= 15)
                    tabWarrior.Label = tabWarrior.Label.ToUpper();
                // set user email
                userEmail = Properties.Settings.Default.UserEmail;
                // create logger instance
                logger = new UsageLogger(officeVersion, userEmail);
                // Check license (still 2014 and e-mail inserted) - disables controls if not valid
                checkLicense();
                // if license invalid, show settings box
                if (!licenseValid)
                    btnAbout_Click(null, null);
                // track statup
                logger.PostUsage("Powerpoint started", null);
                // get style manager
                styles = new StyleLogic();
                // refresh styles list
                refreshStyles();
                // Set event handler to check which buttons should be enabled
                Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        /// <summary>
        /// Checks whether userEmail variable is set. If not, shows settings dialog.
        /// Also checks for year (currently only able to run the toolbar in 2014).
        /// Will wet controls to inactive if license not valid
        /// </summary>
        /// <returns>"License" validity</returns>
        private void checkLicense()
        {
            // if we are in 2015, this version is no logner valid
            if (DateTime.Now.Year > 2014)
            {
                MessageBox.Show("You are using an outdated version of the PowerPoint Warrior.\n" +
                    "Please unistall the current version from \"Programs and Features\" in the Windows Control Panel " +
                    "and visit www.ppwarrior.com for a new version!");
                logger.PostUsage("Tried to load 2014 version");
                licenseValid = false;
            }
            // Check that we have an e-mail
            else if (string.IsNullOrEmpty(userEmail))
            {
                licenseValid = false;
            }
            else
            {
                licenseValid = true;
            }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                // variables: active window and selection descriptor
                var window = Sel.Application.ActiveWindow;
                var selection = new SelectionType();

                // Only check the selection if we are in the slide pane and license valid
                if (window.ActivePane.ViewType == PowerPoint.PpViewType.ppViewSlide && licenseValid)
                {
                    // Essentially exceptions on selection types should not occur, since buttons should be
                    // enable only when the correct selection is made by the user.
                    // Exceptions will be thrown if the buttons are not disabled/enabled correctly, and 
                    // the user tries to perform an action e.g. when there is no shape range in the selection
                    // This makes the following code VERY IMPORTANT to get right!  

                    // One shape
                    selection.ShapesOne = (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                        Sel.Type == PowerPoint.PpSelectionType.ppSelectionText) &&
                        Sel.ShapeRange.Count == 1;
                    // Exactly two shapes
                    selection.ShapesTwo = Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                        Sel.ShapeRange.Count == 2;
                    // More than one shape
                    selection.ShapesMoreThanOne = Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                        Sel.ShapeRange.Count > 1;
                    // Shapes or text in one shape
                    selection.ShapesOrText = Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                        Sel.Type == PowerPoint.PpSelectionType.ppSelectionText;
                    // Text inside a table
                    selection.TableText = Sel.Type == PowerPoint.PpSelectionType.ppSelectionText &&
                        Sel.ShapeRange[1].HasTable == Office.MsoTriState.msoTrue;
                    // One table
                    selection.TableOne = (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                        Sel.Type == PowerPoint.PpSelectionType.ppSelectionText) &&
                        Sel.ShapeRange.Count == 1 &&
                        Sel.ShapeRange[1].HasTable == Office.MsoTriState.msoTrue;

                    // Only do the icon and checked at this point
                    checkSelectionBoxes(Globals.ThisAddIn.Application.ActiveWindow.Selection); 
                }
                else
                {
                    selection.SetAllFalse();
                }

                // These when on one shape or text (i.e. one shape)
                //button.Enabled = selection.ShapesOne;
                // These when one shape or text OR two shapes
                //button.Enabled = selection.ShapesOne || selection.ShapesTwo;
                // These when more than one shape
                btnSameHeight.Enabled = selection.ShapesMoreThanOne;
                btnSameWidth.Enabled = selection.ShapesMoreThanOne;
                // These when exactly 2 shapes
                btnSwapPos.Enabled = selection.ShapesTwo;
                // These when shapes or text (text implicitly means one shape)
                toggleAutoFit.Enabled = selection.ShapesOrText;
                toggleWordWrap.Enabled = selection.ShapesOrText;
                galleryStyles.Enabled = selection.ShapesOrText;
                // These when text in a table (i.e. one cell)
                btnPasteFromExcel.Enabled = selection.TableText;
                // These when one table
                //buttonFormatTable.Enabled = selection.TableOne;
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        #region Internal functions

        private void checkSelectionBoxes(PowerPoint.Selection selection)
        {
            // Counters for shapes w/ word wrap and resize
            double countWordwrap = 0;
            double countResize = 0;
            // Go through all shapes and count how many has word wrap / resize / etc
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionNone && selection.Type != PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    // Check if shape has text
                    bool hasText = shape.HasTextFrame == Office.MsoTriState.msoTrue;
                    // If text, check accordingly
                    if (hasText)
                    {
                        // Count word wrap
                        countWordwrap += shape.TextFrame2.WordWrap == Office.MsoTriState.msoTrue ? 1 : 0;
                        // Count resize, leaving counter as-is if no resize
                        if (shape.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeShapeToFitText)
                            // Add one if normal resize
                            countResize++;
                        else if (shape.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeMixed ||
                            shape.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeTextToFitShape)
                            // Add 0.5 if weird resize
                            countResize = countResize + 0.5;
                    }
                }
                // If counters equal shape count, set normal image and enable
                // if counter > 0 but less than shape count, set grey image and enable
                // if counter = 0, set normal image and disable
                if (countResize == selection.ShapeRange.Count)
                {
                    toggleAutoFit.Checked = true;
                    toggleAutoFit.Image = Properties.Resources.IconResize;
                }
                else if (countResize > 0)
                {
                    toggleAutoFit.Checked = true;
                    toggleAutoFit.Image = Properties.Resources.IconResizeGrey;
                }
                else
                {
                    toggleAutoFit.Checked = false;
                    toggleAutoFit.Image = Properties.Resources.IconResize;
                }
                // The same for word wrap
                if (countWordwrap == selection.ShapeRange.Count)
                {
                    toggleWordWrap.Checked = true;
                    toggleWordWrap.Image = Properties.Resources.IconWordwrap;
                }
                else if (countWordwrap > 0)
                {
                    toggleWordWrap.Checked = true;
                    toggleWordWrap.Image = Properties.Resources.IconWordwrapGrey;
                }
                else
                {
                    toggleWordWrap.Checked = false;
                    toggleWordWrap.Image = Properties.Resources.IconWordwrap;
                }
            }
        }

        private void logUsage(object sender, RibbonControlEventArgs e)
        {
            if (Properties.Settings.Default.EnableLogging)
            {
                // Get label text, reflection from http://stackoverflow.com/questions/1196991/get-property-value-from-string-using-reflection-in-c-sharp
                string action = sender.GetType().GetProperty("Label") != null ?
                    sender.GetType().GetProperty("Label").GetValue(sender, null).ToString() :
                    e.Control.Id;
                // Log usage
                if (logger != null)
                    logger.PostUsage("Feature used", action);
            }
        }

        #endregion

        private void btnPasteFromExcel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Table pptTable = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1].Table;
                ToolsAndFormatting.PasteFromExcel(pptTable);

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void toggleAutoFit_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsAndFormatting.ToggleAutoFit(selection, ((RibbonToggleButton)sender).Checked);
                // After this, we need to re-check the controls
                checkSelectionBoxes(selection);

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void toggleWordWrap_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsAndFormatting.ToggleWordWrap(selection, ((RibbonToggleButton)sender).Checked);
                // After this, we need to re-check the controls
                checkSelectionBoxes(selection);

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void buttonSameHeightOrWidth_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                HeightOrWidth heightOrWidth = e.Control.Id == btnSameHeight.Id ? HeightOrWidth.Height : HeightOrWidth.Width;
                ToolsSizeAndPosition.SameHeightOrWidth(selection, heightOrWidth);

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void btnSwapPos_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsSizeAndPosition.SwapPositions(selection);

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void galleryStyles_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((RibbonButton)sender).Id == btnSaveStyle.Id)
                {
                    PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    if (!((selection.Type == PowerPoint.PpSelectionType.ppSelectionText || selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) &&
                        selection.ShapeRange.Count == 1))
                    {
                        MessageBox.Show("Please select at least one shape to apply style to.");
                        return;
                    }

                    styles.SaveStyle(selection);
                }
                else
                {
                    styles.DeleteStyle(Globals.ThisAddIn.Application);
                }
                refreshStyles();

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void refreshStyles()
        {
            // clear everything
            galleryStyles.Items.Clear();
            // if there are no styles, insert note
            if (styles.Styles == null || styles.Styles.Count == 0)
            {
                RibbonDropDownItem ddi = Factory.CreateRibbonDropDownItem();
                ddi.Label = "(no styles)";
                galleryStyles.Items.Add(ddi);
                return;
            }

            foreach (var style in styles.Styles)
            {
                RibbonDropDownItem ddi = Factory.CreateRibbonDropDownItem();
                ddi.Label = style.Key;
                ddi.Tag = style.Key;
                ddi.OfficeImageId = "CellStylesGallery";
                galleryStyles.Items.Add(ddi); 
            }
        }

        private void galleryStyles_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                string styleName = ((RibbonGallery)sender).SelectedItem.Label;
                // if this is the (no styles) -style, just return and don't apply anything
                if (styleName == "(no styles)")
                    return;
                // apply style
                styles.ApplyStyle(styleName, selection);

                logUsage(sender, e);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get window pointer
                IntPtr pointer = new IntPtr(Globals.ThisAddIn.Application.HWND);
                IWin32Window w = Control.FromHandle(pointer);

                using (FormSettings settings = new FormSettings())
                {
                    // show the dialog
                    settings.ShowDialog(w);
                    // if email now exists, update user identity
                    if(!string.IsNullOrEmpty(Properties.Settings.Default.UserEmail))
                    {
                        userEmail = Properties.Settings.Default.UserEmail;
                        logger.UpdateIdentity(userEmail);
                    }
                }

                // validate license again, since it might have been updated
                checkLicense();
                // re-evaluate which buttons to show, since license may have been updated
                Application_WindowSelectionChange(Globals.ThisAddIn.Application.ActiveWindow.Selection);

                // If this was called from code, assume it was the initial e-mail prompt
                if (sender == null)
                {
                    logger.PostUsage("Showed initial e-mail prompt");
                }
                else
                {
                    logUsage(sender, e);
                }
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void btnFeedback_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Email.SendFeedback("Feedback for PowerPoint Warrior");
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }

        private void btnUpgrade_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start
                        ("http://www.ppwarrior.com/?utm_source=in-app&utm_medium=v1&utm_campaign=upgrade");
            }
            catch (Exception ex)
            {
                Warrior_Common.Exceptions.Handle(ex, officeVersion, userEmail);
            }
        }
    }
}
