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
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using Microsoft.Win32;
using IWshRuntimeLibrary;

namespace Planz
{
    public static class StartProcess
    {
        public const string START_FOLDER = "Projects";
        public const string PLANZ_XML_FILENAME = "planz.xml";
        public const string XOOML_XML_FILENAME = "XooML.xml";
        public const string PLANNER_XML_FILENAME = "planner.xml";
        public const string APPDATA_FOLDERNAME = "Planz";
        public const string TODAY_PLUS = "Today+";
        public const string LIFE = "Life";
        public const string JOURNAL = "Journal";
        public const string INBOX = "Notes";
        public const string GOALS_THINGS = "Goals & Things";
        public const string ROLES_PEOPLE = "Roles & People";
        public const string DAYS_AHEAD = "The Days Ahead";
        public static string ROOT_PATH = System.IO.Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)) +
            System.IO.Path.DirectorySeparatorChar;

        public static string START_PATH = ROOT_PATH + START_FOLDER + System.IO.Path.DirectorySeparatorChar;
        public static string TODAY_PLUS_PATH = START_PATH + TODAY_PLUS + System.IO.Path.DirectorySeparatorChar;
        public static string INBOX_PATH = TODAY_PLUS_PATH + INBOX + System.IO.Path.DirectorySeparatorChar;
        public static string DAYS_AHEAD_PATH = TODAY_PLUS_PATH + DAYS_AHEAD + System.IO.Path.DirectorySeparatorChar;
        public static string START_PLANZ_XML = ROOT_PATH + START_FOLDER + System.IO.Path.DirectorySeparatorChar + XOOML_XML_FILENAME;
        public static string START_PLANZ8_XML = ROOT_PATH + START_FOLDER + System.IO.Path.DirectorySeparatorChar + PLANZ_XML_FILENAME;
        public static string START_PLANNER_XML = ROOT_PATH + START_FOLDER + System.IO.Path.DirectorySeparatorChar + PLANNER_XML_FILENAME;
        public static string APPDATA_PLANZ_PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
            System.IO.Path.DirectorySeparatorChar + APPDATA_FOLDERNAME + System.IO.Path.DirectorySeparatorChar;
        public static string LIFE_PATH = ROOT_PATH + LIFE + System.IO.Path.DirectorySeparatorChar;
        public static string JOURNAL_PATH = LIFE_PATH + JOURNAL + System.IO.Path.DirectorySeparatorChar;
        public static string GOALS_THINGS_PATH = LIFE_PATH + GOALS_THINGS + System.IO.Path.DirectorySeparatorChar;
        public static string ROLES_PEOPLE_PATH = LIFE_PATH + ROLES_PEOPLE + System.IO.Path.DirectorySeparatorChar;   
        public static int MAX_EXTRACTTEXT_LENGTH = 500;
        public static int MAX_EXTRACTNAME_LENGTH = 80;

        private const string REGISTRY_MENU = "Folder\\shell\\NewMenuOption";
        private const string REGISTRY_COMMAND = "Folder\\shell\\NewMenuOption\\command";
        private const string REGISTRY_NAME = "Planz";
        private static string APPLICATION_PATH = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
        public static string PLANZ_VERSION = "8.2";
        public static string BUILD_VERSION = "r1504";
        public static string BUILD_DATE = "06/11/2010";

        public static bool FirstTime = false;

        public static void CheckFirstTimeRunning()
        {
            if (Directory.Exists(START_PATH) == false)
            {
                Directory.CreateDirectory(START_PATH);
            }
            if (System.IO.File.Exists(START_PLANZ_XML) == false &&
                System.IO.File.Exists(START_PLANZ8_XML) == false &&
                System.IO.File.Exists(START_PLANNER_XML) == false)
            {
                FirstTime = true;

                // Create Planz setup XMLs and folders
                CreatePlanzFolders();

                // Set Properties
                Properties.Settings.Default.LastFocusedElementGuid = "32b24fe1-45e3-4965-bfa0-38330dd6864a";
                Properties.Settings.Default.Save();
            }
            if (Directory.Exists(APPDATA_PLANZ_PATH) == false)
            {
                Directory.CreateDirectory(APPDATA_PLANZ_PATH);
            }
        }

        private static void CreatePlanzFolders()
        {
            // Copy Planz.xml under Projects
            FileStream fs = new FileStream(START_PLANZ_XML, FileMode.Create);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.InnerXml = Properties.Resources.Planz_Root;
            xmlDoc.Save(fs);
            fs.Close();

            // Create default folders
            try
            {
                #region Today+

                Directory.CreateDirectory(START_PATH + TODAY_PLUS);

                fs = new FileStream(START_PATH + TODAY_PLUS + System.IO.Path.DirectorySeparatorChar + XOOML_XML_FILENAME, FileMode.Create);
                xmlDoc = new XmlDocument();
                xmlDoc.InnerXml = Properties.Resources.Planz_Today;
                xmlDoc.Save(fs);
                fs.Close();

                #endregion

                #region The Days Ahead

                Directory.CreateDirectory(DAYS_AHEAD_PATH);

                #endregion

                #region Notes

                Directory.CreateDirectory(INBOX_PATH);

                fs = new FileStream(INBOX_PATH + System.IO.Path.DirectorySeparatorChar + XOOML_XML_FILENAME, FileMode.Create);
                xmlDoc = new XmlDocument();
                xmlDoc.InnerXml = Properties.Resources.Planz_Notes;
                xmlDoc.Save(fs);
                fs.Close();

                #endregion

                #region User Manual Link

                WshShellClass wshShell = new WshShellClass();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(START_PATH + "Planner User Manual.lnk");
                shortcut.TargetPath = "http://kftf.ischool.washington.edu/planner/User_Manual/HTML/user_manual.html";
                shortcut.Description = "http://kftf.ischool.washington.edu/planner/User_Manual/HTML/user_manual.html";
                shortcut.Save();

                #endregion

                #region My First Project

                string f_mfp = "My First Project";

                Directory.CreateDirectory(START_PATH + f_mfp);

                fs = new FileStream(START_PATH + f_mfp + System.IO.Path.DirectorySeparatorChar + XOOML_XML_FILENAME, FileMode.Create);
                xmlDoc = new XmlDocument();
                xmlDoc.InnerXml = Properties.Resources.Planz_MyFirstProject;
                xmlDoc.Save(fs);
                fs.Close();

                #endregion

                #region Life

                string lifeFolder = ROOT_PATH + LIFE;
                Directory.CreateDirectory(lifeFolder);

                Directory.CreateDirectory(lifeFolder + System.IO.Path.DirectorySeparatorChar + "Journal");
                Directory.CreateDirectory(lifeFolder + System.IO.Path.DirectorySeparatorChar + "Goals & Things");
                Directory.CreateDirectory(lifeFolder + System.IO.Path.DirectorySeparatorChar + "Roles & People");
                Directory.CreateDirectory(lifeFolder + System.IO.Path.DirectorySeparatorChar + "Labels");
                Directory.CreateDirectory(lifeFolder + System.IO.Path.DirectorySeparatorChar + "Regions");

                JournalControl.CreateJournalFolders(DateTime.Now.Year, JOURNAL_PATH);

                #endregion
             
            }
            catch (Exception)
            {

            }
        }
    }
}
