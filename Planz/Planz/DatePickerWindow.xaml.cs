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
    public enum DatePickerCondition
    {
        PowerDEarlier,
        PowerDLater,
    }

    public partial class DatePickerWindow : Window
    {
        public DatePickerWindow(DatePickerCondition dpc)
        {
            InitializeComponent();

            this.cal.SelectionMode = Microsoft.Windows.Controls.CalendarSelectionMode.SingleDate;

            switch (dpc)
            {
                case DatePickerCondition.PowerDEarlier:
                    this.cal.DisplayDateEnd = DateTime.Now.AddDays(-2);
                    break;
                case DatePickerCondition.PowerDLater:
                    this.cal.DisplayDateStart = DateTime.Now.AddDays(3);
                    break;
                default:
                    break;
            };
        }

        public DateTime SelectedDate
        {
            get;
            set;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            SelectedDate = this.cal.SelectedDate.Value;
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
