using System;
using System.Windows.Forms;

namespace Setup
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			Installer installer = new Installer();
			installer.InstallApplication(@"https://ppwarrior.blob.core.windows.net/install/setup.exe");
			MessageBox.Show("Installer object created.");
		}
	}
}
