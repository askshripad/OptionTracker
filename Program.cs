using System;
using System.Windows.Forms;

namespace NSEDayMarketTracker
{
    /// <summary>
    /// The main program.
    /// </summary>
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MarketTracker());
        }
    }
}