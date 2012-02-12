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

namespace Planz
{
    public static class HeadingNameConverter
    {
        public static string ConvertFromHeadingNameToFolderName(Element element)
        {
            string folderName = element.NoteText.Trim();
            
            if (folderName == String.Empty)
            {
                folderName = element.ID.ToString();
            }

            while (folderName.EndsWith("."))
            {
                folderName = folderName.Substring(0, folderName.Length - 1) + "^d";
            }
            folderName = folderName.Replace("/", "^f");
            folderName = folderName.Replace("\\", "^s");
            folderName = folderName.Replace("*", "^a");
            folderName = folderName.Replace("?", "^q");
            folderName = folderName.Replace("\"", "^b");
            folderName = folderName.Replace("<", "^l");
            folderName = folderName.Replace(">", "^r");
            folderName = folderName.Replace("|", "^v");
            folderName = folderName.Replace(":", "^c");
            folderName = folderName.Replace("\t", "^t");

            return folderName;
        }

        public static string ConvertFromFolderNameToHeadingName(string folderName)
        {
            string headingName = folderName;

            while (headingName.EndsWith("^d"))
            {
                headingName = headingName.Substring(0, headingName.Length - 2) + ".";
            }
            headingName = headingName.Replace("^f", "/");
            headingName = headingName.Replace("^s", "\\");
            headingName = headingName.Replace("^a", "*");
            headingName = headingName.Replace("^q", "?");
            headingName = headingName.Replace("^b", "\"");
            headingName = headingName.Replace("^l", "<");
            headingName = headingName.Replace("^r", ">");
            headingName = headingName.Replace("^v", "|");
            headingName = headingName.Replace("^c", ":");

            return headingName;
        }

        public static bool Exist(string folderPath)
        {
            return System.IO.Directory.Exists(folderPath);
        }

        public static Boolean IsValidGuid(String s)
        {
            try
            {
                Guid newGuid = new Guid(s);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

    }

    public static class FileNameConverter
    {
        public static string ConvertFromNoteTextToFileName(Element element)
        {
            string fileName = element.NoteText.Trim();

            if (fileName == String.Empty)
            {
                fileName = element.ID.ToString();
            }

            while (fileName.EndsWith("."))
            {
                fileName = fileName.Substring(0, fileName.Length - 1);
            }
            fileName = fileName.Replace("/", "");
            fileName = fileName.Replace("\\","");
            fileName = fileName.Replace("*", "");
            fileName = fileName.Replace("?", "");
            fileName = fileName.Replace("\"","");
            fileName = fileName.Replace("<", "");
            fileName = fileName.Replace(">", "");
            fileName = fileName.Replace("|", "");
            fileName = fileName.Replace(":", "");

            return fileName;
        }
    }

    public static class OutlookSupportFunction
    {
        public static string GetObjectDescriptor(System.IO.MemoryStream ms)
        {
            int offset = 52;
            int len = (int)ms.Length - offset;
            if (len < 1)
            {
                return String.Empty;
            }
            else
            {
                byte[] byteArray = new byte[len + 99];
                ms.Position = offset;
                int count = ms.Read(byteArray, 0, len);
                string strBuffer = Encoding.Unicode.GetString(byteArray);
                int nullOffset = strBuffer.IndexOf('\0');
                string strDesc = strBuffer.Substring(0, nullOffset);
                return strDesc;
            }
        }

        public static string GenerateShortcutTargetPath()
        {
            string path = OutlookSupportFunction.GetOutlookApplicationPath() + "OUTLOOK.EXE";
            return path;
        }

        private static string GetOutlookApplicationPath()
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE", false);

            return (string)key.GetValue("Path");
        }

        public static string GetOutlookEntryIDFromShortcut(string shortcutPath)
        {
            string entryID = String.Empty;

            IWshRuntimeLibrary.WshShellClass wshShell = new IWshRuntimeLibrary.WshShellClass();
            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(shortcutPath);
            if (shortcut != null)
            {
                string arguments = shortcut.Arguments;

                // Sample arguments = outlook:000000007E9E0115FB90D348B60A3CCC64FCEAC364002000

                entryID = arguments.Substring(arguments.IndexOf("outlook:") + 8);
            }

            return entryID;
        }
    }

    public class SearchFunction
    {
        private List<Element> targetList = new List<Element>();
        private int index = 0;

        public SearchFunction()
        {

        }

        public List<Element> TargetList
        {
            get
            {
                return this.targetList;
            }
            set
            {
                this.targetList = value;
            }
        }

        public void Clear()
        {
            this.targetList.Clear();
            this.index = 0;
        }

        public Element GetNextElement()
        {
            if (targetList.Count == 0)
            {
                return null;
            }
            else if (targetList.Count <= index)
            {
                index = 0;
                return targetList[index++];
            }
            else
            {
                return targetList[index++];
            }
        }
    }
}
