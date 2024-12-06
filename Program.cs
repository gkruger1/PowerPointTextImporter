using System;
using System.Windows.Forms;

namespace PowerPointTextImporter
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Generate the icon before starting the application
            IconResource.GenerateIcon();

            Application.Run(new Form1());
        }
    }
}