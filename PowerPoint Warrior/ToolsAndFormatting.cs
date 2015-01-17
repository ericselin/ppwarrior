using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPoint_Warrior
{
    class ToolsAndFormatting
    {
        public static void PasteFromExcel(PowerPoint.Table pptTable)
        {
            string clipboardText = Clipboard.GetText();
            char[] newlineChar = { '\n' };
            char[] tabChar = { '\t' };
            int selectedTableRow = 0;
            int selectedTableCol = 0;

            // Get the currently selected row and column
            for (int iRow = 1; iRow <= pptTable.Rows.Count; iRow++)
            {
                for (int iCol = 1; iCol <= pptTable.Columns.Count; iCol++)
                {
                    if (pptTable.Cell(iRow, iCol).Selected)
                    {
                        selectedTableRow = iRow;
                        selectedTableCol = iCol;
                        // Return from the nested loop
                        goto nextStep;
                    }
                }
            }

        nextStep:

            // Take out '\r' and extra newline char from clipboar data
            clipboardText = clipboardText.Replace("\r", "").TrimEnd(newlineChar);

            // Get number of rows and cols in clipboard
            int clipboardRowCount = clipboardText.Split(newlineChar).Count();
            int clipboardColCount = clipboardText.Split(newlineChar)[0].Split(tabChar).Count();

            // Check that the data will fit into the table given the currently selected cell
            if (selectedTableRow + clipboardRowCount - 1 > pptTable.Rows.Count ||
                selectedTableCol + clipboardColCount - 1 > pptTable.Columns.Count)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Your copied data would not fit into this table.\nPlease make table bigger or select another cell.");
                return;
            }

            // Create the multidimensional data array for clipboard data
            //string[,] clipboardTable = new string[clipboardRowCount, clipboardColCount];

            // Insert clipboard data into the table
            string[] clipboardRows = clipboardText.Split(newlineChar);
            string[] clipboardCells;
            for (int iRow = 0; iRow < clipboardRowCount; iRow++)
            {
                clipboardCells = clipboardRows[iRow].Split(tabChar);
                for (int iCol = 0; iCol < clipboardColCount; iCol++)
                {
                    pptTable.Cell(selectedTableRow + iRow, selectedTableCol + iCol).Shape.TextFrame.TextRange.Text =
                        clipboardCells[iCol];
                }
            }
        }

        public static void FormatBullets(PowerPoint.Selection selection)
        {
            // Select the text range either from whole shape or text element
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    formatBullets(shape.TextFrame2.TextRange);
                }
            else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                formatBullets(selection.TextRange2);
        }

        private static void formatBullets(Office.TextRange2 textRange)
        {
            // Loop through the paragraphs if paragraphs exists  (user selected many characters)
            if (textRange.Paragraphs.Count > 0)
            {
                foreach (Office.TextRange2 paragraph in textRange.Paragraphs)
                {
                    formatBullets(paragraph.ParagraphFormat);
                }
            }
            // If no selection (only cursor position), only adjust that paragraph
            else
            {
                formatBullets(textRange.ParagraphFormat);
            }
        }

        private static void formatBullets(Office.ParagraphFormat2 paragraphFormat)
        {
            // Set the font to default (Arial), because minus looks like a long dash in Courier (default)
            paragraphFormat.Bullet.UseTextFont = Office.MsoTriState.msoTrue;
            paragraphFormat.Bullet.Font.Name = "Arial";
            if (paragraphFormat.IndentLevel < 2)
                // Totally not sure where the int code for the bullet is from, I think I debugged it to find out
                paragraphFormat.Bullet.Character = 8226;
            else
            {
                // This just parses the "-" to a char, and then to an int, debugging shows it is int value 45
                paragraphFormat.Bullet.Character = (int)Char.Parse("-");
            }
            // Also change the indentation, note that with small fonts this will look dumb
            paragraphFormat.FirstLineIndent = -0.5f * Constants.PointsPerCm;
            paragraphFormat.LeftIndent = 0.5f * Constants.PointsPerCm * paragraphFormat.IndentLevel;
        }

        public static void LineBelow(PowerPoint.DocumentWindow window)
        {
            foreach (PowerPoint.Shape shape in window.Selection.ShapeRange)
            {
                // make line and fill "no color"
                shape.Fill.Visible = Office.MsoTriState.msoFalse;
                shape.Line.Visible = Office.MsoTriState.msoFalse;
                
                window.View.Slide.Shapes.AddLine(
                    shape.Left, shape.Top + shape.Height, shape.Left + shape.Width, shape.Top + shape.Height);
            }
        }

        public static void RemoveEffects(PowerPoint.Selection selection)
        {
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                // Remove reflection (for some reason this needs to be before everything else)
                shape.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone;
                // Remove glow
                shape.Glow.Radius = 0f;
                shape.Glow.Transparency = 1.0f;
                // Remove soft edges
                shape.SoftEdge.Type = Office.MsoSoftEdgeType.msoSoftEdgeTypeNone;
                // Remove bevel and 3d rotation
                shape.ThreeD.Visible = Office.MsoTriState.msoFalse;
                // Remove shadow (for some reason this needs to be last)
                shape.Shadow.Visible = Office.MsoTriState.msoFalse;
                // Remove text shadow
                shape.TextFrame.TextRange.Font.Shadow = Office.MsoTriState.msoFalse;
            }
        }

        public static void SetLanguage(PowerPoint.Slides slides, Office.MsoLanguageID language)
        {
            foreach (PowerPoint.Slide slide in slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    // go to helper method to set the language of the shape
                    setLanguageOfShape(shape, language);
                }
            }
            // Also set the default language of the presentation
            Globals.ThisAddIn.Application.ActivePresentation.DefaultLanguageID = language;
        }

        private static void setLanguageOfShape(PowerPoint.Shape shape, Office.MsoLanguageID language)
        {
            // If this has a text frame, set the language of that
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.LanguageID = language;
            }
            // if it is a table, set language of each cell
            else if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Row row in shape.Table.Rows)
                {
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        cell.Shape.TextFrame.TextRange.LanguageID = language;
                    }
                }
            }
            // if it is a group, start traversing the grouped items tree
            else if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape groupedShape in shape.GroupItems)
                {
                    setLanguageOfShape(groupedShape, language);
                }
            }
        }

        public static void FormatTable(PowerPoint.Table pptTable)
        {
            // Style doc from http://code.msdn.microsoft.com/office/PowerPoint-2010-Interact-ea2fbe1b
            // No fill, table grid. Tested only with PowerPoint 2010
            pptTable.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}", false);
            // Set borders to 3/4
            foreach (PowerPoint.Row row in pptTable.Rows)
            {
                foreach (PowerPoint.Cell cell in row.Cells)
                {
                    // All borders
                    cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 0.75f;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 0.75f;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 0.75f;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 0.75f;
                }
            }
        }

        public static void ToggleAutoFit(PowerPoint.Selection selection, bool disable)
        {
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame2.AutoSize = disable ? Office.MsoAutoSize.msoAutoSizeShapeToFitText : Office.MsoAutoSize.msoAutoSizeNone;
                }
            }
        }

        public static void ToggleWordWrap(PowerPoint.Selection selection, bool disable)
        {
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame2.WordWrap = disable ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                }
            }
        }

        internal static void RemoveNotes(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (slide.HasNotesPage == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.Shape notes = slide.NotesPage.Shapes.Placeholders[2];
                    if (notes.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        notes.TextFrame.DeleteText();
                        notes.TextFrame2.DeleteText();
                    }
                }
            }
        }

        internal static void RemoveAnimations(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                // Remove transitions
                slide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                // Remove effects of individual shapes
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    shape.AnimationSettings.Animate = Office.MsoTriState.msoFalse;
                }
            }
        }
    }
}
