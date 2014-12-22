using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPoint_Warrior
{
    enum Edition
    {
        Trial, Free, Starter, Premium, Corporate
    }

    // Create the selection struct
    public struct SelectionType
    {
        // One shape
        public bool ShapesOne,
        // Exactly two shapes
        ShapesTwo,
        // More than one shape
        ShapesMoreThanOne,
        // Shapes or text in one shape
        ShapesOrText,
        // Text inside a table
        TableText,
        // One table
        TableOne,
        // Show if valid
        Valid;
        // Set all as false
        public void SetAllFalse()
        {
            ShapesOne = ShapesTwo = ShapesMoreThanOne = ShapesOrText = TableText = TableOne = Valid = false;
        }
    };

    public static class Constants
    {
        public static float PointsPerCm = 28.3464566929134f;
    }

    public class PowerPointPosition
    {
        public float Left, Top, Width, Height;
    }

    public enum HeightOrWidth
    {
        Height, Width
    }

    public enum TopOrLeft
    {
        Top, Left
    }

    public enum SelectSimilarTypes
    {
        SelectSimilarColorLine, 
        SelectSimilarColor, 
        SelectSimilarLine, 
        SelectSimilarHeight, 
        SelectSimilarWidth, 
        SelectSimilarHorizontal, 
        SelectSimilarVertical
    }
}
