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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Planz
{
    public partial class FileNameWindow : Window
    {
        private string fileName = String.Empty;
        private string filePath = null;
        private string fileExt = null;

        private bool isFolder = false;

        private ICCAssociationType type;

        public FileNameWindow(string fileFullName, ICCAssociationType type)
        {
            InitializeComponent();

            this.filePath = System.IO.Path.GetDirectoryName(fileFullName) + System.IO.Path.DirectorySeparatorChar;
            this.fileExt = System.IO.Path.GetExtension(fileFullName);
            this.TextBoxFileName.Text = System.IO.Path.GetFileNameWithoutExtension(fileFullName);

            this.type = type;
            if (type == ICCAssociationType.OutlookEmailMessage)
            {
                this.Label.Content = "Please enter the Email subject:";
            }

            this.TextBoxFileName.SelectionStart = this.TextBoxFileName.Text.Length;
            this.TextBoxFileName.Focus();
        }

        public FileNameWindow(string fileFullName, string label)
        {
            InitializeComponent();

            if ((isFolder = System.IO.Directory.Exists(fileFullName)) == true)
            {
                this.filePath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(fileFullName)) + System.IO.Path.DirectorySeparatorChar;
                this.fileExt = String.Empty;
                this.TextBoxFileName.Text = System.IO.Path.GetFileNameWithoutExtension(System.IO.Path.GetDirectoryName(fileFullName));
            }
            else
            {
                this.filePath = System.IO.Path.GetDirectoryName(fileFullName) + System.IO.Path.DirectorySeparatorChar;
                this.fileExt = System.IO.Path.GetExtension(fileFullName);
                this.TextBoxFileName.Text = System.IO.Path.GetFileNameWithoutExtension(fileFullName);
            }

            this.type = ICCAssociationType.WordDocument;

            this.Label.Content = label;
            this.TextBoxFileName.SelectionStart = this.TextBoxFileName.Text.Length;
            this.TextBoxFileName.Focus();
        }

        public string FileName
        {
            get
            {
                return fileName; 
            }
            set
            {
                fileName = value;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            fileName = this.TextBoxFileName.Text;

            if (this.type == ICCAssociationType.OutlookEmailMessage)
            {
                DialogResult = true;
            }
            else if (isFolder)
            {
                string folderPath = filePath + fileName;

                if (System.IO.Directory.Exists(folderPath))
                {
                    System.Windows.MessageBox.Show("There is already a folder with the same name in this location.",
                        "Invalid Folder Name");
                }
                else if (!FileNameChecker.HasInvalidChar(fileName))
                {
                    string invalidChar = String.Empty;
                    foreach (string s in FileNameChecker.InvalidCharList)
                    {
                        invalidChar += s + " ";
                    }
                    System.Windows.MessageBox.Show("A folder name cannot contain any of the following characters:\n"
                        + invalidChar, "Invalid Folder Name");
                }
                else if (FileNameChecker.IsLongName(fileName))
                {
                    System.Windows.MessageBox.Show("Folder name is too long.",
                        "Invalid Folder Name");
                }
                else if (fileName == String.Empty)
                {
                    System.Windows.MessageBox.Show("You must type a folder name.",
                        "Invalid Folder Name");
                }
                else
                {
                    DialogResult = true;
                }
            }
            else
            {
                string fileFullPath = filePath + fileName + fileExt;

                if (System.IO.File.Exists(fileFullPath))
                {
                    System.Windows.MessageBox.Show("There is already a file with the same name in this location.",
                        "Invalid File Name");
                }
                else if (!FileNameChecker.HasInvalidChar(fileName))
                {
                    string invalidChar = String.Empty;
                    foreach (string s in FileNameChecker.InvalidCharList)
                    {
                        invalidChar += s + " ";
                    }
                    System.Windows.MessageBox.Show("A file name cannot contain any of the following characters:\n"
                        + invalidChar, "Invalid File Name");
                }
                else if (FileNameChecker.IsLongName(fileName))
                {
                    System.Windows.MessageBox.Show("File name is too long.",
                        "Invalid File Name");
                }
                else if (fileName == String.Empty)
                {
                    System.Windows.MessageBox.Show("You must type a file name.",
                        "Invalid File Name");
                }
                else
                {
                    DialogResult = true;
                }
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
