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
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Text.RegularExpressions;
using IWshRuntimeLibrary;

namespace Planz
{
    public static class FileTypeHandler
    {
        private static string appDataFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + 
            "\\Planz\\";
        private const string fileIconPath = "Images\\Icons\\";
        private const string templatePath = "Templates\\";
        
        private const string IconFormat = ".png";
        private const string folderIcon = ".folder";
        private const string folderShortcutIcon = ".folderShortcut";
        private const string webIcon = ".url";
        private const string mailIcon = ".mail";
        private const string appointmentIcon = ".appointment";
        private const string twitterIcon = ".twitter";
        private const string missingIcon = ".missing";
        private const string shortcutMark = "Shortcut";

        private const string word = "template.docx";
        private const string excel = "template.xlsx";
        private const string ppt = "template.pptx";
        private const string onenote = "template.one";
        private const string notepad = "template.txt";

        public static string GetIcon(ElementAssociationType type, string fileFullName)
        {
            switch (type)
            {
                case ElementAssociationType.File:
                    return GetFileIcon(fileFullName);
                case ElementAssociationType.FileShortcut:
                    return GetFileShortcutIcon(fileFullName);
                case ElementAssociationType.Folder:
                    return GetFolderIcon();
                case ElementAssociationType.FolderShortcut:
                    return GetFolderShortcutIcon();
                case ElementAssociationType.Web:
                    return GetWebShortcutIcon();
                case ElementAssociationType.Tweet:
                    return GetTweetShortcutIcon();
                case ElementAssociationType.Email:
                    return GetEmailShortcutIcon();
                case ElementAssociationType.Appointment:
                    return GetAppointmentShortcutIcon();
                default:
                    return String.Empty;
            };
        }

        private static string GetFileIcon(string fileFullName)
        {
            try
            {
                string ext = Path.GetExtension(fileFullName).ToLower();
                string copyPath = appDataFolderPath + fileIconPath + ext + IconFormat;

                if (System.IO.File.Exists(copyPath))
                {

                }
                else
                {
                    Icon icon = Icon.ExtractAssociatedIcon(fileFullName);
                    Bitmap bitmap = icon.ToBitmap();
                    MemoryStream strm = new MemoryStream();
                    bitmap.Save(strm, System.Drawing.Imaging.ImageFormat.Png);
                    bitmap.Save(copyPath);
                }

                return copyPath;
            }
            catch (Exception)
            {
                return String.Empty;
            }
        }

        private static string GetFileShortcutIcon(string fileFullName)
        {
            try
            {
                WshShellClass wshShell = new WshShellClass();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(fileFullName);

                string targetFileFullName = shortcut.TargetPath;
                
                string ext = Path.GetExtension(targetFileFullName).ToLower();
                string copyPath = appDataFolderPath + fileIconPath + ext + shortcutMark + IconFormat;

                if (System.IO.File.Exists(copyPath))
                {

                }
                else
                {
                    Icon icon = Icon.ExtractAssociatedIcon(fileFullName);
                    Bitmap bitmap = icon.ToBitmap();
                    MemoryStream strm = new MemoryStream();
                    bitmap.Save(strm, System.Drawing.Imaging.ImageFormat.Png);
                    bitmap.Save(copyPath);
                }
                
                return copyPath;
            }
            catch (Exception)
            {
                return String.Empty;
            }
        }

        private static string GetFolderIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + folderIcon + IconFormat;

            return copyPath;
        }

        private static string GetFolderShortcutIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + folderShortcutIcon + IconFormat;

            return copyPath;
        }

        private static string GetWebShortcutIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + webIcon + IconFormat;

            return copyPath;
        }

        private static string GetEmailShortcutIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + mailIcon + IconFormat;

            return copyPath;
        }

        private static string GetAppointmentShortcutIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + appointmentIcon + IconFormat;

            return copyPath;
        }

        private static string GetTweetShortcutIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + twitterIcon + IconFormat;

            return copyPath;
        }

        public static string GetMissingFileIcon()
        {
            string copyPath = appDataFolderPath + fileIconPath + missingIcon + IconFormat;

            return copyPath;
        }

        public static int GetFileType(string fileFullName)
        {
            ElementAssociationType eat = ElementAssociationType.File;
            string ext = System.IO.Path.GetExtension(fileFullName).ToLower();
            switch (ext)
            {
                case ".lnk":
                    try
                    {
                        WshShellClass wshShell = new WshShellClass();
                        IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(fileFullName);
                        string targetPath = shortcut.TargetPath;
                        if (Directory.Exists(targetPath))
                        {
                            eat = ElementAssociationType.FolderShortcut;
                        }
                        else
                        {
                            eat = ElementAssociationType.FileShortcut;
                        }
                    }
                    catch (Exception)
                    {
                        eat = ElementAssociationType.FileShortcut;
                    }
                    break;
                default:
                    eat = ElementAssociationType.File;
                    break;
            };
            return (int)eat;
        }

        public static string GetTemplate(ICCAssociationType at)
        {
            switch (at)
            {
                case ICCAssociationType.WordDocument:
                    return appDataFolderPath + templatePath + word;
                case ICCAssociationType.ExcelSpreadsheet:
                    return appDataFolderPath + templatePath + excel;
                case ICCAssociationType.PowerPointPresentation:
                    return appDataFolderPath + templatePath + ppt;
                case ICCAssociationType.OneNoteSection:
                    return appDataFolderPath + templatePath + onenote;
                case ICCAssociationType.NotepadTextDocument:
                    return appDataFolderPath + templatePath + notepad;
                default:
                    return null;
            };
        }
    }

    public static class ShortcutNameConverter
    {
        public const string fileLinkExt = ".lnk";
        public const string webLinkExt = ".url";
        public const string mailLinkExt = ".lnk";

        // fileName: NOT filePath
        public static string GenerateShortcutNameFromFileName(string fileName, string folderPath)
        {
            string shortcutName = RenameShortcutName(System.IO.Path.GetFileNameWithoutExtension(fileName), fileLinkExt, folderPath);

            return shortcutName;
        }

        public static string GenerateShortcutNameFromWebTitle(string webTitle, string folderPath)
        {
            string shortcutName = RenameShortcutName(webTitle, fileLinkExt, folderPath);

            return shortcutName;
        }

        public static string GenerateShortcutNameFromEmailSubject(string emailSubject, string folderPath)
        {
            string shortcutName = RenameShortcutName(emailSubject, fileLinkExt, folderPath);

            return shortcutName;
        }

        public static string GenerateShortcutNameFromAppointmentSubject(string appointSubject, string folderPath)
        {
            string shortcutName = RenameShortcutName(appointSubject, fileLinkExt, folderPath);

            return shortcutName;
        }

        public static string RenameShortcutName(string shortcutNameWithoutExt, string ext, string folderPath)
        {
            string shortcutName = FileNameChecker.Rename(shortcutNameWithoutExt, ext, folderPath);

            return shortcutName;
        }
    }

    public static class FileNameChecker
    {
        public static string[] InvalidCharList = new string[]{ @"\", "/", ":", "*", "?", "<", ">", "|", "\"" };
 
        public static bool HasInvalidChar(string fileNameWithoutExt)
        {
            foreach (string s in InvalidCharList.ToList())
            {
                if (fileNameWithoutExt.Contains(s))
                {
                    return false;
                }
            }
            return true;
        }

        private const int MaxFileNameLength = 255;
        private const int MaxFolderNameLength = 80;//247;

        public static bool IsLongName(string fileNameWithoutExt)
        {
            if (fileNameWithoutExt.Length > MaxFileNameLength)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsLongNameFolder(string folderName)
        {
            if (folderName.Length > MaxFolderNameLength)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool Exist(string fileFullPath)
        {
            return System.IO.File.Exists(fileFullPath);
        }

        public static string Rename(string fileNameWithoutExt, string ext, string folderPath)
        {
            string newFileNameWithoutExt = String.Empty;

            foreach (char s in fileNameWithoutExt)
            {
                if (!InvalidCharList.Contains(s.ToString()))
                {
                    newFileNameWithoutExt += s;
                }
            }

            if (IsLongName(newFileNameWithoutExt))
            {
                newFileNameWithoutExt = newFileNameWithoutExt.Substring(0, MaxFileNameLength);
            }

            while (newFileNameWithoutExt.EndsWith("."))
                newFileNameWithoutExt = newFileNameWithoutExt.TrimEnd('.');

            string newFileFullPath = folderPath + newFileNameWithoutExt + ext;

            string newFileNameWithoutExtWithoutNumber = newFileNameWithoutExt;
            Regex regex = new Regex(@"\w \(\[\d\]\)");
            if (regex.Match(newFileNameWithoutExtWithoutNumber).Success)
            {
                int end = newFileNameWithoutExtWithoutNumber.LastIndexOf("(");
                newFileNameWithoutExtWithoutNumber = newFileNameWithoutExtWithoutNumber.Substring(0, end - 1);
            }
            int index = 2;
            while (Exist(newFileFullPath))
            {
                newFileFullPath = folderPath + newFileNameWithoutExtWithoutNumber + " (" + index.ToString() + ")" + ext;
                index++;
            }

            string newFileName = System.IO.Path.GetFileName(newFileFullPath);

            return newFileName;
        }
    }

    public static class ICCFileNameHandler
    {
        private const string ext_txt = ".txt";
        private const string ext_word = ".docx";
        private const string ext_excel = ".xlsx";
        private const string ext_ppt = ".pptx";
        private const string ext_onenote = ".one";
        private const string ext_email = ".lnk";
        private const string ext_web = ".lnk";


        public static string GenerateFileName(string fileName, ICCAssociationType type)
        {
            string newFileName = String.Empty;
            foreach (char ch in fileName)
            {
                if (!FileNameChecker.InvalidCharList.Contains(ch.ToString()))
                {
                    newFileName += ch;
                }
            }

            switch(type)
            {
                case ICCAssociationType.NotepadTextDocument:
                    return newFileName + ext_txt;
                case ICCAssociationType.WordDocument:
                    return newFileName + ext_word;
                case ICCAssociationType.ExcelSpreadsheet:
                    return newFileName + ext_excel;
                case ICCAssociationType.PowerPointPresentation:
                    return newFileName + ext_ppt;
                case ICCAssociationType.OutlookEmailMessage:
                    if (newFileName == "")
                    {
                        newFileName = "New Email";
                    }
                    return newFileName + ext_email;
                case ICCAssociationType.OneNoteSection:
                    return newFileName + ext_onenote;
                case ICCAssociationType.TwitterUpdate:
                    return newFileName + ext_web;
                default:
                    return String.Empty;
            };
        }
    }
}
