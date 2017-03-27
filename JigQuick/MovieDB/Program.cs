using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace JigQuick
{
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

            Temp tmp = new Temp();
            string formType = tmp.readIni("APPLICATION BEHAVIOR", "FORM TYPE");

            switch (formType)
            {
                case "DOCKING":
                    Application.Run(new frmDocking());
                    break;
                case "TIPTAIL":
                    Application.Run(new frmTipTail());
                    break;
                case "MULTI":
                    Application.Run(new frmMulti());
                    break;
                case "TRIPLE":
                    Application.Run(new frmTriple());
                    break;
                case "OMNI":
                    Application.Run(new frmOmni());
                    break;
                case "NOJIG99":
                    Application.Run(new frmNoJig99());
                    break;
                case "WITHJIG99":
                    Application.Run(new frmWithJig99());
                    break;
                case "COILASSY99":
                    Application.Run(new frmCoilAssy99());
                    break;
                default:
                    Application.Run(new frmDocking());
                    break;
            }
        }
    }
}