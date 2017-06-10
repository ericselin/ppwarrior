using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPoint_Warrior
{
	public static class Exceptions
	{
		public static void Handle(Exception ex, bool showMessage = true)
		{
			// show friendly message 
			if (showMessage)
				System.Windows.Forms.MessageBox.Show(
					"Unfortunately, an error occured in the Warrior add-in!\n" +
					"We have logged the error and will try to adress it as soon as possible.\n" +
					"In the mean time, if you have any questions or comments, please e-mail eric.selin@gmail.com");
		}
	}
}
