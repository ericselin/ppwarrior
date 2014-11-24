using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPoint_Warrior
{
    public static class ToolsSizeAndPosition
    {
        public static void SameHeightOrWidth(PowerPoint.Selection selection, HeightOrWidth heightOrWidth)
        {
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                if (heightOrWidth == HeightOrWidth.Height)
                    shape.Height = selection.ShapeRange[1].Height;
                else
                    shape.Width = selection.ShapeRange[1].Width;
            }
        }

        public static PowerPointPosition PickUpPosition(PowerPoint.Selection selection)
        {
            PowerPointPosition point = new PowerPointPosition();
            point.Left = selection.ShapeRange[1].Left;
            point.Top = selection.ShapeRange[1].Top;
            point.Width = selection.ShapeRange[1].Width;
            point.Height = selection.ShapeRange[1].Height;
            return point;
        }

        public static void ApplyPosition(PowerPoint.Selection selection, PowerPointPosition point)
        {
            selection.ShapeRange[1].Left = point.Left;
            selection.ShapeRange[1].Top = point.Top;
            selection.ShapeRange[1].Width = point.Width;
            selection.ShapeRange[1].Height = point.Height;
        }

        public static void SwapPositions(PowerPoint.Selection selection)
        {
            float left = selection.ShapeRange[1].Left;
            float top = selection.ShapeRange[1].Top;
            selection.ShapeRange[1].Left = selection.ShapeRange[2].Left;
            selection.ShapeRange[1].Top = selection.ShapeRange[2].Top;
            selection.ShapeRange[2].Left = left;
            selection.ShapeRange[2].Top = top;
        }

        public static void SplitObject(PowerPoint.Selection selection)
        {
            PowerPoint.Shape shape = selection.ShapeRange[1];
            // Get number of paragraphs and original height
            int n = shape.TextFrame2.TextRange.Paragraphs.Count;
            // Set height to 1 / [paragraph count] of original
            shape.Height = shape.Height / n;
            // For each additional paragraph: duplicate, set left and top, remove unnecessary paragraphs
            for (int i = 2; i <= n; i++)
            {
                // Duplicate original shape to create new shape
                PowerPoint.Shape s = shape.Duplicate()[1];
                s.Left = shape.Left;
                s.Top = shape.Top + (i-1) * shape.Height;
                // Go through paragraphs in new shape and only keep the one we are interested in
                for (int p = n; p > 0; p--)
                {
                    if (i != p)
                    {
                        s.TextFrame2.TextRange.Paragraphs[p].Delete();
                    }
                }
                // Remove line break
                trimNewline(s.TextFrame2.TextRange.Paragraphs[1]);
            }
            // Remove paragraphs 2 - n from original
            for (int p = 2; p <= n; p++)
            {
                shape.TextFrame2.TextRange.Paragraphs[2].Delete();
            }
            // Remove last empty line from original shape
            trimNewline(shape.TextFrame2.TextRange.Paragraphs[1]);
        }

        private static void trimNewline(Office.TextRange2 paragraph)
        {
            Office.TextRange2 lastChar = paragraph.Characters[paragraph.Characters.Count];
            if (lastChar.Text == "\r")
            {
                paragraph.Characters[paragraph.Characters.Count].Delete();
            }
        }

        internal static void AlignTopToBottom(PowerPoint.Selection selection, TopOrLeft topLeft)
        {
            PowerPoint.Shape anchor;
            var list = ToolsCommon.getEdgeShapeAndChildren(selection.ShapeRange, topLeft, out anchor);

            foreach (var s in list)
            {
                if (topLeft == TopOrLeft.Top)
                    s.Top = anchor.Top + anchor.Height;
                else if (topLeft == TopOrLeft.Left)
                    s.Left = anchor.Left + anchor.Width;
            }
        }
    }
}
