using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPoint_Warrior
{
    public static class ToolsCommon
    {
        public static string getClickOnceLocation()
        {
            // Get resource path - from http://robindotnet.wordpress.com/2010/07/11/how-do-i-programmatically-find-the-deployed-files-for-a-vsto-add-in/
            // Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            // CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = System.IO.Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            return ClickOnceLocation;
        }

        public static bool getSelection(PowerPoint.PpSelectionType selectionType, out PowerPoint.Selection selection)
        {
            PowerPoint.Selection _selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (_selection.Type == selectionType)
            {
                selection = _selection;
                return true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(
                    String.Format("You have selected {0}.\nPlease select {1} instead!",
                    getSelectionName(_selection.Type), getSelectionName(selectionType)));
                selection = _selection;
                return false;
            }

        }

        public static string getSelectionName(PowerPoint.PpSelectionType selectionType)
        {
            switch (selectionType)
            {
                case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionNone:
                    return "nothing";
                case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes:
                    return "shapes";
                case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides:
                    return "slides";
                case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText:
                    return "text";
                default:
                    return "undefined";
            }
        }

        public static List<PowerPoint.Shape> getEdgeShapeAndChildren(PowerPoint.ShapeRange shapes, TopOrLeft edge, out PowerPoint.Shape anchor)
        {
            var list = new List<PowerPoint.Shape>();

            // create initial list with all shapes
            foreach (PowerPoint.Shape s in shapes)
            {
                list.Add(s);
            }

            PowerPoint.Shape edgeShape = null;

            // get the edge shape
            switch (edge)
            {
                case TopOrLeft.Left:
                    edgeShape = (from s in list
                                 orderby s.Left
                                 select s).FirstOrDefault();
                    break;
                case TopOrLeft.Top:
                    edgeShape = (from s in list
                                 orderby s.Top
                                 select s).FirstOrDefault();
                    break;
                default:
                    break;
            }

            // remove the edge shape from the list
            list.Remove(edgeShape);

            anchor = edgeShape;
            return list;
        }
    }
}
