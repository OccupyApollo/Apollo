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
    /// <summary>
    /// Interaction logic for DeleteWindow.xaml
    /// </summary>
    public partial class DeleteWindow : Window
    {
        private string message = String.Empty;

        public DeleteWindow(DeleteMessageType type)
        {
            InitializeComponent();

            switch (type)
            {
                case DeleteMessageType.File:
                    ChangeButtonStyle(ButtonStyle.YESNO);
                    message = "Are you sure you want to move this file to the Recycle Bin?";
                    break;
                case DeleteMessageType.Shortcut:
                    ChangeButtonStyle(ButtonStyle.OKCANCEL);
                    message = "This shortcut will be moved to the Recycle Bin.";
                    break;
                case DeleteMessageType.HeadingWithChildren:
                    ChangeButtonStyle(ButtonStyle.YESNO);
                    message = "Are you sure you want to delete this heading and move its associated folder and all content to the Recycle Bin?";
                    break;
                case DeleteMessageType.HeadingWithoutChildren:
                    // no message
                    break;
                case DeleteMessageType.InplaceExpansionHeading:
                    ChangeButtonStyle(ButtonStyle.OKCANCEL);
                    message = "This heading will be removed and its associated shortcut will be moved to the Recycle Bin.";
                    break;
                case DeleteMessageType.NoteWithFileAssociation:
                    ChangeButtonStyle(ButtonStyle.YESNO);
                    message = "Are you sure you want to delete this note and move its associated file to the Recycle Bin?";
                    break;
                case DeleteMessageType.NoteWithShortcutAssociation:
                    ChangeButtonStyle(ButtonStyle.OKCANCEL);
                    message = "This note will be removed and its associated shortcut will be moved to the Recycle Bin.";
                    break;
                case DeleteMessageType.NoteWithoutAssociation:
                    // no messgae
                    break;
                case DeleteMessageType.MultipleItems:
                    message = "Are you sure you want to delete all these notes?";
                    break;
                case DeleteMessageType.Default:
                    ChangeButtonStyle(ButtonStyle.YESNO);
                    message = "Are you sure you want to delete this heading/note?";
                    break;
            };
            this.Message.Text = message;
        }

        private void YES_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void NO_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ChangeButtonStyle(ButtonStyle style)
        {
            switch (style)
            {
                case ButtonStyle.OKCANCEL:
                    this.BtnYes.Content = "OK";
                    this.BtnNo.Content = "Cancel";
                    this.BtnYes.IsDefault = true;
                    this.BtnNo.IsDefault = false;
                    break;
                case ButtonStyle.YESNO:
                    this.BtnYes.Content = "Yes";
                    this.BtnNo.Content = "No";
                    this.BtnYes.IsDefault = false;
                    this.BtnNo.IsDefault = true;
                    break;
            };
        }
    }

}
