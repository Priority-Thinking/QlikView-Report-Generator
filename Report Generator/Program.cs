using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Report_Generator
{
    static class Program
    {
        /// <summary>
        /// Project Name: Report Generator
        /// Project Description: Qlikview to MS Word Automated Reporting Tool
        /// Company: Institute 4 Priority Thinking, LLC.
        /// Contact Website: www.prioritythinking.com
        /// Authors: Tim Kendrick, John Murray, Mitali Ajgaonkar, Grant Parker
        /// Date: 7/11/2016
        /// License: GNU LESSER GENERAL PUBLIC LICENSE VERSION 3 (GPLv3)
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new GeneratorSpace.GenerateReport());
        }
    }
}
