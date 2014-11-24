using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPoint_Warrior
{
    public static class ToolsSelection
    {
        public static void GoToSlide(PowerPoint.View view, string slideNumberString)
        {
            int slideNumber;
            if (Int32.TryParse(slideNumberString, out slideNumber))
            {
                try
                {
                    view.GotoSlide(slideNumber);
                }
                catch (Exception)
                {
                    System.Windows.Forms.MessageBox.Show(String.Format("Could not load slide number {0}.\nTry again.", slideNumber));
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Could not parse number.\nTry again.");
            }
        }

        public static void SelectSimilar(PowerPoint.Selection selection, SelectSimilarTypes selectType)
        {
            PowerPoint.Shapes slideShapes = selection.SlideRange.Shapes;
            PowerPoint.Shape originalShape = selection.ShapeRange[1];

            // Clear selection
            selection.Unselect();

            switch (selectType)
            {
                case SelectSimilarTypes.SelectSimilarColorLine:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Type != Office.MsoShapeType.msoTable &&
                            shape.Fill.Visible == originalShape.Fill.Visible &&
                            (shape.Fill.ForeColor.RGB == originalShape.Fill.ForeColor.RGB || shape.Fill.Visible == Office.MsoTriState.msoFalse) &&
                            shape.Line.DashStyle == originalShape.Line.DashStyle &&
                            shape.Line.Weight == originalShape.Line.Weight &&
                            shape.Line.ForeColor.RGB == originalShape.Line.ForeColor.RGB)
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                case SelectSimilarTypes.SelectSimilarColor:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Fill.Visible == originalShape.Fill.Visible &&
                            (shape.Fill.ForeColor.RGB == originalShape.Fill.ForeColor.RGB || shape.Fill.Visible == Office.MsoTriState.msoFalse))
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                case SelectSimilarTypes.SelectSimilarLine:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Type != Office.MsoShapeType.msoTable &&
                            shape.Line.DashStyle == originalShape.Line.DashStyle &&
                            shape.Line.Weight == originalShape.Line.Weight &&
                            shape.Line.ForeColor.RGB == originalShape.Line.ForeColor.RGB)
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                case SelectSimilarTypes.SelectSimilarHeight:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Height > originalShape.Height * 0.9 &&
                            shape.Height < originalShape.Height * 1.1)
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                case SelectSimilarTypes.SelectSimilarWidth:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Width > originalShape.Width * 0.9 &&
                            shape.Width < originalShape.Width * 1.1)
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                case SelectSimilarTypes.SelectSimilarHorizontal:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Top > originalShape.Top - 15 &&
                            shape.Top < originalShape.Top + 15)
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                case SelectSimilarTypes.SelectSimilarVertical:
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Left > originalShape.Left - 15 &&
                            shape.Left < originalShape.Left + 15)
                        {
                            // Select the shape
                            shape.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    break;
                default:
                    break;
            }
        }
    }
}
