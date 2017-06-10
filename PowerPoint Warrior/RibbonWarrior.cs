using System;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace PowerPoint_Warrior
{
    public partial class RibbonWarrior
    {
        // position used by pick up / apply pos.
        PowerPointPosition position;

        private void RibbonWarrior_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                // Set event handler to check which buttons should be enabled
                Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
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
                if (window.ActivePane.ViewType == PowerPoint.PpViewType.ppViewSlide)
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
                btnApplyPosition.Enabled = selection.ShapesOne && position != null;
                btnPickUpPosition.Enabled = selection.ShapesOne;
                gallerySelectSimilar.Enabled = selection.ShapesOne;
                btnSplitShape.Enabled = selection.ShapesOne;
                // These when one shape or text OR two shapes
                btnHeaderLine.Enabled = selection.ShapesOne || selection.ShapesTwo;
                // These when more than one shape
                btnSameHeight.Enabled = selection.ShapesMoreThanOne;
                btnSameWidth.Enabled = selection.ShapesMoreThanOne;
                galleryAlign.Enabled = selection.ShapesMoreThanOne;
                // These when exactly 2 shapes
                btnSwapPos.Enabled = selection.ShapesTwo;
                // These when shapes or text (text implicitly means one shape)
                toggleAutoFit.Enabled = selection.ShapesOrText;
                toggleWordWrap.Enabled = selection.ShapesOrText;
                btnLineBelow.Enabled = selection.ShapesOrText;
                btnFormatBullets.Enabled = selection.ShapesOrText;
                btnRemoveEffects.Enabled = selection.ShapesOrText;
                // These when text in a table (i.e. one cell)
                btnPasteFromExcel.Enabled = selection.TableText;
                // These when one table
                btnFormatTable.Enabled = selection.TableOne;

				#region Enable all buttons
				// Uncomment when taking screenshots
				//btnApplyPosition.Enabled = licenseValid;
				//btnPickUpPosition.Enabled = licenseValid;
				//gallerySelectSimilar.Enabled = licenseValid;
				//btnSplitShape.Enabled = licenseValid;
				//btnHeaderLine.Enabled = licenseValid;
				//btnSameHeight.Enabled = licenseValid;
				//btnSameWidth.Enabled = licenseValid;
				//galleryAlign.Enabled = licenseValid;
				//btnSwapPos.Enabled = licenseValid;
				//toggleAutoFit.Enabled = licenseValid;
				//toggleWordWrap.Enabled = licenseValid;
				//galleryStyles.Enabled = licenseValid;
				//btnLineBelow.Enabled = licenseValid;
				//btnFormatBullets.Enabled = licenseValid;
				//btnRemoveEffects.Enabled = licenseValid;
				//btnPasteFromExcel.Enabled = licenseValid;
				//btnFormatTable.Enabled = licenseValid;
				#endregion
			}
			catch (Exception ex)
            {
                Exceptions.Handle(ex, false);
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

        #endregion

        private void btnPasteFromExcel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Table pptTable = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1].Table;
                ToolsAndFormatting.PasteFromExcel(pptTable);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
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
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
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
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void buttonSameHeightOrWidth_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                HeightOrWidth heightOrWidth = e.Control.Id == btnSameHeight.Id ? HeightOrWidth.Height : HeightOrWidth.Width;
                ToolsSizeAndPosition.SameHeightOrWidth(selection, heightOrWidth);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnSwapPos_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsSizeAndPosition.SwapPositions(selection);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
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
                Exceptions.Handle(ex);
            }
        }

        private void btnRemoveEffects_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsAndFormatting.RemoveEffects(selection);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnLineBelow_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;
                ToolsAndFormatting.LineBelow(window);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnFormatBullets_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsAndFormatting.FormatBullets(selection);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnFormatTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Table pptTable = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1].Table;
                ToolsAndFormatting.FormatTable(pptTable);

            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnSetLanguage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (e.Control.Id == btnSetLanguageEnglish.Id)
                    ToolsAndFormatting.SetLanguage(Globals.ThisAddIn.Application.ActivePresentation.Slides, Office.MsoLanguageID.msoLanguageIDEnglishUS);
                else if (e.Control.Id == btnSetLanguageFinnsh.Id)
                    ToolsAndFormatting.SetLanguage(Globals.ThisAddIn.Application.ActivePresentation.Slides, Office.MsoLanguageID.msoLanguageIDFinnish);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void gallerySelectSimilar_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                SelectSimilarTypes selectType;

                if (e.Control.Id == buttonSelectSimilarColorLine.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarColorLine;
                }
                else if (e.Control.Id == buttonSelectSimilarColor.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarColor;
                }
                else if (e.Control.Id == buttonSelectSimilarLine.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarLine;
                }
                else if (e.Control.Id == buttonSelectSimilarHeight.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarHeight;
                }
                else if (e.Control.Id == buttonSelectSimilarWidth.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarWidth;
                }
                else if (e.Control.Id == buttonSelectSimilarHorizontal.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarHorizontal;
                }
                else if (e.Control.Id == buttonSelectSimilarVertical.Id)
                {
                    selectType = SelectSimilarTypes.SelectSimilarVertical;
                }
                else
                    return;

                ToolsSelection.SelectSimilar(selection, selectType);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void galleryAlign_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                TopOrLeft topLeft;

                if (e.Control.Id == buttonAlignLeftToRight.Id)
                    topLeft = TopOrLeft.Left;
                else if (e.Control.Id == buttonAlignTopToBottom.Id)
                    topLeft = TopOrLeft.Top;
                else
                    return;

                ToolsSizeAndPosition.AlignTopToBottom(selection, topLeft);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnPickUpPosition_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                position = ToolsSizeAndPosition.PickUpPosition(selection);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnApplyPosition_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsSizeAndPosition.ApplyPosition(selection, position);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnSplitShape_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                ToolsSizeAndPosition.SplitObject(selection);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void editBoxGoToSlide_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.View view = Globals.ThisAddIn.Application.ActiveWindow.View;
                string slideNumberString = editBoxGoToSlide.Text;
                ToolsSelection.GoToSlide(view, slideNumberString);

                editBoxGoToSlide.Text = "";
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnHeaderLine_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.DocumentWindow window = Globals.ThisAddIn.Application.ActiveWindow;

                ToolsGuidelines.HeaderLine(window);
            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnRemoveNotes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                ToolsAndFormatting.RemoveNotes(presentation);

            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }

        private void btnRemoveAnimations_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                ToolsAndFormatting.RemoveAnimations(presentation);

            }
            catch (Exception ex)
            {
                Exceptions.Handle(ex);
            }
        }
    }
}
