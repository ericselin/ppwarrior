using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;


namespace Style_Manager
{
    [Serializable()]
    public class Style
    {
        public Style()
        {
            Fill = new StyleModels.Fill();
            Line = new StyleModels.Line();
            Font = new StyleModels.Font();
        }

        public virtual StyleModels.Fill Fill { get; set; }
        public virtual StyleModels.Line Line { get; set; }
        public virtual StyleModels.Font Font { get; set; }
    }
}


namespace Style_Manager.StyleModels
{
    [Serializable()]
    public class Style
    {
        public bool Enabled { get; set; }
    }

    [Serializable()]
    public class Color
    {
        public Color(PowerPoint.ColorFormat pptColor)
        {
            ObjectThemeColor = pptColor.ObjectThemeColor;
            RGB = pptColor.RGB;
        }
        public Color() {}

        public int RGB { get; set; }
        public Office.MsoThemeColorIndex ObjectThemeColor { get; set; }
    }

    [Serializable()]
    public class Fill : Style
    {
        public Fill(PowerPoint.FillFormat pptFill)
        {
            FillColor = new Color(pptFill.ForeColor);
            Visible = pptFill.Visible;
            Enabled = true;
        }
        public Fill()
        { 
            Enabled = false;
            FillColor = new Color();
        }

        public Color FillColor { get; set; }
        public Office.MsoTriState Visible { get; set; }
    }

    [Serializable()]
    public class Line : Style
    {
        public Line(PowerPoint.LineFormat pptLine)
        {
            LineStyle = pptLine.Style;
            DashStyle = pptLine.DashStyle;
            ForeColor = new Color(pptLine.ForeColor);
            Weight = pptLine.Weight;
            Visible = pptLine.Visible;
            Enabled = true;
        }
        public Line()
        {
            Enabled = false;
            ForeColor = new Color();
        }

        public Office.MsoLineStyle LineStyle { get; set; }
        public Office.MsoLineDashStyle DashStyle { get; set; }
        public Color ForeColor { get; set; }
        public float Weight { get; set; }
        public Office.MsoTriState Visible { get; set; }
    }

    [Serializable()]
    public class Font : Style
    {
        public Font(PowerPoint.Font pptFont)
        {
            Bold = pptFont.Bold;
            Italic = pptFont.Italic;
            Size = pptFont.Size;
            Name = pptFont.Name;
            FontColor = new Color(pptFont.Color);
            Enabled = true;
        }
        public Font() 
        {
            FontColor = new Color();
            Enabled = false; 
        }

        public Office.MsoTriState Bold { get; set; }
        public Office.MsoTriState Italic { get; set; }
        public float Size { get; set; }
        public string Name { get; set; }
        public Color FontColor { get; set; }
    }
}


