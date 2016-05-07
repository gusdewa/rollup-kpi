using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace rollup_kpi
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            IndicatorRollup ind = new IndicatorRollup();
            string rootUrl = "https://mcai4.sharepoint.com/sites/ims";
            try
            {
                ind.GetIndicatorWithRollup(rootUrl + "/pm");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                ind.GetIndicatorWithRollup(rootUrl + "/gp");
            }
            catch { }
            try
            {
                ind.GetIndicatorWithRollup(rootUrl + "/hn");
            }
            catch { }
        }
    }
}
