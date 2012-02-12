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
    /// Interaction logic for TwitterWindow.xaml
    /// </summary>
    public partial class TwitterWindow : Window
    {
        private string username = String.Empty;
        private string password = String.Empty;
        private string tweet = String.Empty;

        private const int TWEET_MAX_CHAR = 140;

        public TwitterWindow(string defaultText)
        {
            InitializeComponent();

            this.TextBoxTweet.Text = defaultText;
            this.TextBoxTweet.SelectionStart = defaultText.Length;
            this.TextBoxTweet.Focus();
            if (Properties.Settings.Default.TwitterUsername != String.Empty)
            {
                this.TextBoxUsername.Text = Properties.Settings.Default.TwitterUsername;
                this.TextBoxPassword.Password = Properties.Settings.Default.TwitterPassword;
                this.CheckBoxRememberMe.IsChecked = true;
            }
        }

        public string Username
        {
            get { return this.username; }
        }

        public string Password
        {
            get { return this.password; }
        }

        public string Tweet
        {
            get { return this.tweet; }
        }

        private void Update_Click(object sender, RoutedEventArgs e)
        {
            this.tweet = this.TextBoxTweet.Text;
            if (this.tweet.Length > TWEET_MAX_CHAR)
            {
                MessageBox.Show("A tweet cannot be longer than 140 characters.");
            }
            else
            {
                this.username = this.TextBoxUsername.Text;
                this.password = this.TextBoxPassword.Password;
                DialogResult = true;
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void CheckBoxRememberMe_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.TwitterUsername = this.TextBoxUsername.Text;
            Properties.Settings.Default.TwitterPassword = this.TextBoxPassword.Password;
            Properties.Settings.Default.Save();
        }

        private void CheckBoxRememberMe_Unchecked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.TwitterUsername = String.Empty;
            Properties.Settings.Default.TwitterPassword = String.Empty;
            Properties.Settings.Default.Save();
        }
    }
}
