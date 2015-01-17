using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Warrior_Common;
using System.Windows.Forms;

namespace PowerPoint_Warrior
{
    class ToolsGuidelines
    {
        internal static void InsertFooter(Microsoft.Office.Interop.PowerPoint.DocumentWindow window)
        {
            // see if footer exists
            PowerPoint.Shape footer = null;
            bool footerExists;
            try
            {
                footer = window.View.Slide.Shapes["August Footer"];
                footerExists = true;
                // if user wants to format exising footer, continue, otherwise return
                DialogResult result = MessageBox.Show(
                    "Footer already created.\nDo you want to format the existing footer instead?",
                    "", System.Windows.Forms.MessageBoxButtons.YesNo);
                if (result == DialogResult.No)
                {
                    return;
                }
            }
            catch (Exception)
            {
                footerExists = false;
            }
            // if no footer, create a new footer
            if (!footerExists)
            {
                // add the text box
                // from http://msdn.microsoft.com/en-us/library/ff743980(v=office.14).aspx
                // AddTextbox(Orientation, Left, Top, Width, Height)
                footer = window.View.Slide.Shapes.
                    AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 150, 100, 400, 100);
                // set text and name
                footer.TextFrame2.TextRange.Text = "[footer text]";
                footer.Name = "August Footer";
            }
            // if we came this far, format the footer
            // align bottom and left
            footer.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom;
            footer.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft;
            // autosize and wordwrap
            footer.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText;
            footer.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;
            // font size (should be 8pt) and italics etc
            footer.TextFrame2.TextRange.Font.Size = 8F;
            footer.TextFrame2.TextRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue;
            footer.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
            // margins
            footer.TextFrame2.MarginLeft = 0.1f * Constants.PointsPerCm;
            footer.TextFrame2.MarginRight = 0.1f * Constants.PointsPerCm;
            footer.TextFrame2.MarginTop = 0f;
            footer.TextFrame2.MarginBottom = 0f;
            // no fill / line
            footer.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            footer.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            // set position and width
            footer.Left = 1.1f * Constants.PointsPerCm;
            footer.Top = 17.92f * Constants.PointsPerCm - footer.Height; // align bottom to 17.92, i.e. top at that pos minus height
            footer.Width = 23.2f * Constants.PointsPerCm;
        }

        internal static void HeaderLine(PowerPoint.DocumentWindow window)
        {
            // If one shape selected, add connector line to it
            if (window.Selection.ShapeRange.Count == 1)
            {
                // get the shape and its position
                PowerPoint.Shape shape = window.Selection.ShapeRange[1];
                // create the line 
                createHeaderLine(window, shape);
            }
            // If two shapes, align the connector to the shape
            else
            {
                // Determine which shape is the bottom box (/shape)
                var shapes = window.Selection.ShapeRange;
                var bottomIndex = shapes[1].Top > shapes[2].Top ? 1 : 2;
                var bottom = shapes[bottomIndex];
                var top = bottomIndex == 1 ? shapes[2] : shapes[1];
                // newly created connector to align
                PowerPoint.Shape conn;
                // if top shape is not a connector, create a new connector based on the top shape
                if (top.Connector == Office.MsoTriState.msoFalse)
                {
                    // create the line (inside try to be able to return)
                    conn = createHeaderLine(window, top);
                }
                // else connector is obviously the top shape
                else
                {
                    conn = top;
                }
                // Set the width (and height) of the connector aligning it
                var width = bottom.Left + bottom.Width - conn.Left;
                if (width > 0)
                {
                    conn.Width = width;
                    conn.ConnectorFormat.EndDisconnect();
                    conn.Height = 0f;
                }
                else
                {
                    MessageBox.Show("Cannot align header line, because top shape overlaps bottom shape on the right");
                    return;
                }
            }
        }

        private static PowerPoint.Shape createHeaderLine(PowerPoint.DocumentWindow window, PowerPoint.Shape shape)
        {
            PowerPoint.Shape conn = null;
            try
            {
                // starting x and y values
                float x = shape.Left + shape.Width;
                float y = shape.Top + shape.Height / 2;
                // add the connector line
                // from http://msdn.microsoft.com/en-us/library/ff744679(v=office.14).aspx
                // AddConnector(Type, BeginX, BeginY, EndX, EndY)
                conn = window.View.Slide.Shapes.
                    AddConnector(Office.MsoConnectorType.msoConnectorStraight, x, y, x + 100, y);
                // connect it to the shape
                conn.ConnectorFormat.BeginConnect(shape, 4);
                // set height to 0 (need to make sure end is disconnected first)
                conn.ConnectorFormat.EndDisconnect();
                conn.Height = 0f;
            }
            catch (Exception ex)
            {
                if (conn != null)
                {
                    conn.Delete();
                }
                throw ex;
            }
            // return the created header line
            return conn;
        }
    }
}
