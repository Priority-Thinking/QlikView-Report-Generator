using System;
using System.Windows.Forms;

/*
    This file is part of Report Generator.

    Report Generator is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Report Generator is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Report Generator.  If not, see <http://www.gnu.org/licenses/>.
 */

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
        /// Date: 05/06/2017
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
