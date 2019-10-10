using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenSaver
{
    static class Program
    {
        static void ShowScreenSaver()
        {
            //S'il y a au moins 1 écran
            foreach (Screen screen in Screen.AllScreens)
            {
                ScreenSaverForm screensaver = new ScreenSaverForm(screen.Bounds);
                screensaver.Show();
            }
        }


        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {     
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            ShowScreenSaver();
            Application.Run();          

            System.Diagnostics.Process.GetCurrentProcess().Kill();
            Application.Exit();
        }
    }
}