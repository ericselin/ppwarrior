using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Deployment.Application;

namespace PowerPoint_Warrior
{
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
        TableOne;
        // Set all as false
        public void SetAllFalse()
        {
            ShapesOne = ShapesTwo = ShapesMoreThanOne = ShapesOrText = TableText = TableOne = false;
        }
    };

    public static class Constants
    {
        public const float PointsPerCm = 28.3464566929134f;
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

	public static class Information
	{
		public static string GetUserName()
		{
			return Environment.MachineName + "\\" + Environment.UserName;
		}

		public static string GetAssemblyVersion()
		{
			return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
		}

		public static string GetClickOnceVersion()
		{
			if (ApplicationDeployment.IsNetworkDeployed)
			{
				return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
			}
			return "Debug";
		}
	}
}
