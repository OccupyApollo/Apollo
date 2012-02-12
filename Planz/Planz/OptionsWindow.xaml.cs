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
    public partial class OptionsWindow : Window
    {
        public bool isHeadingFontChanged = false;
        public bool isHeadingSizeChanged = false;
        public bool isNoteFontChanged = false;
        public bool isNoteSizeChanged = false;

        public string headingFont;
        public int headingSize;
        public string noteFont;
        public int noteSize;

        public OptionsWindow()
        {
            InitializeComponent();

            #region Load Font/Size

            foreach (FontFamily ff in Fonts.SystemFontFamilies)
            {
                ComboBoxItem rcbi = new ComboBoxItem();
                rcbi.Content = ff.Source;
                rcbi.FontFamily = new FontFamily(ff.Source);
                rcbi.FontSize = 15;
                this.HeadingLevel1FontFamily.Items.Add(rcbi);
                if (Properties.Settings.Default.HeadingLevel1FontFamily == ff.Source)
                {
                    this.HeadingLevel1FontFamily.SelectedItem = rcbi;
                }
            }
            foreach (FontFamily ff in Fonts.SystemFontFamilies)
            {
                ComboBoxItem rcbi = new ComboBoxItem();
                rcbi.Content = ff.Source;
                rcbi.FontFamily = new FontFamily(ff.Source);
                rcbi.FontSize = 15;
                this.NoteFontFamily.Items.Add(rcbi);
                if (Properties.Settings.Default.NoteFontFamily == ff.Source)
                {
                    this.NoteFontFamily.SelectedItem = rcbi;
                }
            }
            List<int> fontSizeList = new List<int>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36 };
            foreach (int s in fontSizeList)
            {
                ComboBoxItem rcbi = new ComboBoxItem();
                rcbi.Content = s;
                rcbi.FontSize = s;
                this.HeadingLevel1FontSize.Items.Add(rcbi);
                if (Properties.Settings.Default.HeadingLevel1FontSize == s)
                {
                    this.HeadingLevel1FontSize.SelectedItem = rcbi;
                }
            }
            foreach (int s in fontSizeList)
            {
                ComboBoxItem rcbi = new ComboBoxItem();
                rcbi.Content = s;
                rcbi.FontSize = s;
                this.NoteFontSize.Items.Add(rcbi);
                if (Properties.Settings.Default.NoteFontSize == s)
                {
                    this.NoteFontSize.SelectedItem = rcbi;
                }
            }

            #endregion
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void HeadingLevel1FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            headingFont = ((ComboBoxItem)this.HeadingLevel1FontFamily.SelectedItem).Content.ToString();
            if (headingFont != Properties.Settings.Default.HeadingLevel1FontFamily)
            {
                isHeadingFontChanged = true; 
            }
            else
            {
                isHeadingFontChanged = false;
            }
        }

        private void HeadingLevel1FontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            headingSize = Int32.Parse(((ComboBoxItem)this.HeadingLevel1FontSize.SelectedItem).Content.ToString());
            if (headingSize != Properties.Settings.Default.HeadingLevel1FontSize)
            {
                isHeadingSizeChanged = true;
            }
            else
            {
                isHeadingSizeChanged = false;
            }
        }

        private void NoteFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            noteFont = ((ComboBoxItem)this.NoteFontFamily.SelectedItem).Content.ToString();
            if (noteFont != Properties.Settings.Default.NoteFontFamily)
            {
                isNoteFontChanged = true;
            }
            else
            {
                isNoteFontChanged = false;
            }
        }

        private void NoteFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            noteSize = Int32.Parse(((ComboBoxItem)this.NoteFontSize.SelectedItem).Content.ToString());
            if (noteSize != Properties.Settings.Default.NoteFontSize)
            {
                isNoteSizeChanged = true;
            }
            else
            {
                isNoteSizeChanged = false;
            }
        }
    }
}
