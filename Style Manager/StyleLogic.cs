using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace Style_Manager
{
    public class StyleLogic
    {
        public SortedDictionary<string, Style> Styles 
        {
            get; private set;
        }

        public StyleLogic()
        {
            Styles = StyleData.GetStyles();
            if (Styles == null)
                Styles = new SortedDictionary<string, Style>();
        }

        public void SaveStyle(PowerPoint.Selection selection)
        {
            if (!((selection.Type == PowerPoint.PpSelectionType.ppSelectionText || selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) &&
                selection.ShapeRange.Count == 1))
            {
                System.Windows.Forms.MessageBox.Show("Please select one shape to copy attributes from.");
            }
            else
            {
                var shape = selection.ShapeRange[1];

                // save the current style
                var style = new Style();

                style.Fill = new StyleModels.Fill(shape.Fill);
                style.Line = new StyleModels.Line(shape.Line);
                style.Font = new StyleModels.Font(shape.TextFrame.TextRange.Font);

                // save style
                // Get input for placeholder text and set the date
                IntPtr pointer = new IntPtr(selection.Application.HWND);
                IWin32Window w = Control.FromHandle(pointer);
                Style_Manager.SaveDialog saveDialog = new Style_Manager.SaveDialog();
                // refresh teh combo list
                saveDialog.comboStyleName.Items.Clear();
                saveDialog.comboStyleName.Items.AddRange(Styles.Keys.ToArray());
                // show dialog
                if (saveDialog.ShowDialog(w) == DialogResult.OK)
                {
                    string name = saveDialog.comboStyleName.Text;
                    if (name == "")
                    {
                        System.Windows.Forms.MessageBox.Show("Please type a name for a new style or select an existing style to overwrite.");
                        return;
                    }
                    if (!Styles.ContainsKey(name) && Styles.Count >= 5)
                    {
                        System.Windows.Forms.MessageBox.Show("This version supports a maximum of five styles.");
                        return;
                    }
                    Styles[name] = style;
                    StyleData.SaveStyles(Styles); 
                }
                saveDialog.Dispose();
            }
        }

        public void DeleteStyle(PowerPoint.Application application)
        {
            // save style
            // Get input for placeholder text and set the date
            IntPtr pointer = new IntPtr(application.HWND);
            IWin32Window w = Control.FromHandle(pointer);
            // create dialog box
            using (Style_Manager.SaveDialog deleteDialog = new Style_Manager.SaveDialog())
            {
                // refresh the combo list and set for delete mode
                deleteDialog.comboStyleName.Items.Clear();
                deleteDialog.comboStyleName.Items.AddRange(Styles.Keys.ToArray());
                deleteDialog.comboStyleName.DropDownStyle = ComboBoxStyle.DropDownList;
                deleteDialog.lblInstructions.Text = "(select style to delete)";
                // show dialog
                if (deleteDialog.ShowDialog(w) == DialogResult.OK)
                {
                    string name = deleteDialog.comboStyleName.Text;
                    if (name == "")
                    {
                        System.Windows.Forms.MessageBox.Show("Please select an existing style to delete.");
                        return;
                    }
                    Styles.Remove(name);
                    StyleData.SaveStyles(Styles);
                } 
            }
        }

        public void ApplyStyle(string styleName, PowerPoint.Selection selection)
        {
            if ((selection.Type == PowerPoint.PpSelectionType.ppSelectionText || selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) &&
                selection.ShapeRange.Count > 0)
            {
                Style style = Styles[styleName];

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    // fill color
                    if (style.Fill.Enabled)
                        if (style.Fill.Visible == Office.MsoTriState.msoTrue)
                        {
                            if (style.Fill.FillColor.ObjectThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor)
                                shape.Fill.ForeColor.ObjectThemeColor = style.Fill.FillColor.ObjectThemeColor;
                            else
                                shape.Fill.ForeColor.RGB = style.Fill.FillColor.RGB;
                        }
                        else
                            shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    // text
                    if (style.Font.Enabled)
                    {
                        shape.TextFrame.TextRange.Font.Bold = style.Font.Bold;
                        shape.TextFrame.TextRange.Font.Italic = style.Font.Italic;
                        shape.TextFrame.TextRange.Font.Size = style.Font.Size;
                        shape.TextFrame.TextRange.Font.Name = style.Font.Name;
                        if (style.Font.FontColor.ObjectThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor)
                            shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = style.Font.FontColor.ObjectThemeColor;
                        else
                            shape.TextFrame.TextRange.Font.Color.RGB = style.Font.FontColor.RGB;
                    }
                    // line
                    if (style.Line.Enabled)
                    {
                        // first check if there is a line at all
                        if (style.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.DashStyle = style.Line.DashStyle;
                            shape.Line.Style = style.Line.LineStyle;
                            shape.Line.Weight = style.Line.Weight;
                            if (style.Line.ForeColor.ObjectThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor)
                                shape.Line.ForeColor.ObjectThemeColor = style.Line.ForeColor.ObjectThemeColor;
                            else
                                shape.Line.ForeColor.RGB = style.Line.ForeColor.RGB;
                        }
                        else
                            shape.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                }
            }
            else
                System.Windows.Forms.MessageBox.Show("Please select at least one shape to apply style to.");
        }
    }
}
