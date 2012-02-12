/*************************************************************************
*
* DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
* 
* Copyright (c) 2010 by the Keeping Found Things Found group, 
*                       the Information School, University of Washington
*
* Planz - Bring it together: Capture, Connect, … Complete!
*
* This file is part of Planz
*
* Planz is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License version 3
* only, as published by the Free Software Foundation.
*
* Planz is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License version 3 for more details
* (a copy is included in the LICENSE file that accompanied this code).
*
* You should have received a copy of the GNU General Public License
* version 3 along with Planz.  If not, see
* <http://www.gnu.org/licenses/gpl.html>
* for a copy of the GPLv3 License.
*
************************************************************************/

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.IO;

namespace Planz
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Check if another Planz instance exists
            //try
            //{
            //    System.Diagnostics.Process planz_proc = System.Diagnostics.Process.GetProcessesByName(ProcessList.Planz.ToString()).FirstOrDefault();
            //    if (planz_proc != null)
            //    {
            //        Microsoft.VisualBasic.Interaction.AppActivate(planz_proc.Id);
            //        return;
            //    }
            //}
            //catch (Exception)
            //{

            //}


            MainWindow mainWindow;
            if(e.Args.Length>0 && Directory.Exists(e.Args[0].ToString()))
            {
                mainWindow = new MainWindow(e.Args[0].ToString() + System.IO.Path.DirectorySeparatorChar);
            }else
            {
                mainWindow = new MainWindow();
            }

            mainWindow.Show();
        }
    }
}
